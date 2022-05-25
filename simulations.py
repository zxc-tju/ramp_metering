import win32com.client as com
import numpy as np
import pandas as pd


class Simulator:

    def __init__(self):

        # ------ static simulation settings
        self.number_of_step = 3000
        self.sim_break_at = 480
        self.end_of_simulation = 3500
        self.delta_t = 1  # simulation resolution
        self.data_coll_interval = 4

        # ------ simulator initialize
        self.current_time = self.sim_break_at
        self.vissim_interface = None
        self.all_time_all_veh_trj = None
        self.previous_ds_occ = None
        self.previous_ramp_volume = None
        self.control_label = 0
        self.time_interval_id = 0
        self.global_trj_write_label = 0
        self.total_trj_record_num = 3000000
        self.link_length_map = {}
        self.route_links_set = {}
        self.amber_time = []
        self.red_time = []
        self.downstream_occ_collection = []
        self.ramp_flow_collection = []

        headings_list = [['downstream_occ-' + str(i + 1),
                          'ramp_flow-' + str(i + 1),
                          'cal_ramp_volume-' + str(i + 1),
                          'green_time-' + str(i + 1)] for i in range(5)]
        headings = np.array(headings_list).reshape([-1])
        # for save data in each control interval
        self.rm_data = pd.DataFrame(np.zeros((self.number_of_step, 4)), columns=headings)
        # for save vehicle trajectories
        self.all_veh_attributes = np.zeros((10001, 19))
        # 0 global veh id
        # 1 simulation second
        # 2 veh type
        # 3 route decision id
        # 4 route id
        # 5 desired lane
        # 6 destination lane
        # 7 lead target number
        # 8 following distance (Distance to the interaction vehicle)
        # 9 space headway (Distance to the preceding vehicle)
        # 10 lane change state (Direction of the current lane change)
        # 11 acc
        # 12 speed
        # 13 link longitudinal pos
        # 14 link id
        # 15 lane id
        # 16 virtual lane id
        # 17 global_pos  # relative to the zero point (merging point)

        # ------ ramp metering basic parameters
        # TODO for WSH: check following info. for the competition scenario
        self.ramp_num = 5  # number of ramps
        self.saturated_flow = 1800  # single lane saturated flow
        self.cycle_length = 40
        self.virtual_lane_id_map = {(1, 1): 1, (10000, 1): 1, (2, 1): 1, (10001, 1): 1,
                                    (3, 1): 1, (3, 2): 1, (10003, 1): 1,
                                    (10002, 1): 1, (4, 1): 1, (2, 2): 2, (10001, 2): 2,
                                    (3, 3): 2, (2, 2): 2, (10002, 2): 2,
                                    (4, 2): 2, (2, 3): 3, (2, 3): 3, (10001, 3): 3,
                                    (3, 4): 3, (10002, 3): 3, (4, 3): 3,
                                    (2, 4): 4, (10001, 4): 4, (3, 5): 4,
                                    (10002, 4): 4, (4, 4): 4}
        self.zeros_pos_route_coord = {'11': 268.35, '21': 339.458}

    def simulate_scenario(self):
        """
        main process of the simulation
        """

        "==== Connecting the COM Server, Open a new Vissim Window ===="
        self.vissim_interface = com.Dispatch("Vissim.Vissim-32.700")

        "==== Collect Basic Info. of the Simulating Scenario ===="
        self.collect_scenario_data()

        "==== Initialize Vissim and Warmup ===="
        self.init_vissim()
        self.vissim_interface.Simulation.RunContinuous()
        # get vissim state after warmup
        self.get_all_veh_att()

        "==== Start Ramp-metering ===="
        for step in range(self.number_of_step):
            if step == 0:
                # before the ramp metering process keep the ramp signal green
                # note that once signal state is changed, it will keep until simulation ends or changed once again
                self.set_signal(target_key_list='all', state_list=['GREEN'])

            # ------ stop update vehicle's position
            self.vissim_interface.SuspendUpdateGUI()

            # ------ update clock
            self.current_time += step * self.delta_t

            # ------ get and save trajectories at each step
            self.get_all_veh_att()
            if self.global_trj_write_label <= self.total_trj_record_num - 1:
                start_label = self.global_trj_write_label
                self.all_time_all_veh_trj[start_label:start_label + len(self.all_veh_attributes[:, 0]), :] \
                    = self.all_veh_attributes
                self.global_trj_write_label += len(self.all_veh_attributes[:, 0])

            # ------ get and save flow state for each data_coll_interval (4 seconds)
            if self.current_time % self.data_coll_interval == 0:

                # TODO for WSH: we need flow info. of all five ramps using 'get loop state()'
                downstream_occ_list, ramp_flow_list = self.get_loop_state()

                # store new data into the collection
                self.downstream_occ_collection.append(downstream_occ_list)
                self.ramp_flow_collection.append(ramp_flow_list)

                # delete old data, keep only 10 interval in the collection
                if len(self.downstream_occ_collection) == 10:
                    del self.downstream_occ_collection[0]
                    del self.ramp_flow_collection[0]

            "Beginning of a control interval"
            if self.current_time % self.cycle_length == 0:

                # ------ update control interval id
                self.control_label = self.control_label + 1

                # ------ get time interval id
                self.time_interval_id = int(self.current_time / self.cycle_length)

                for i in range(self.ramp_num):
                    # ------ save info. from vissim
                    downstream_occ_collection_array = np.array(self.downstream_occ_collection)
                    ramp_flow_collection_array = np.array(self.ramp_flow_collection)

                    self.rm_data.loc[self.control_label,
                                     'downstream_occ' + str(i + 1)] = np.mean(downstream_occ_collection_array[:, i])
                    self.rm_data.loc[self.control_label,
                                     'ramp_flow' + str(i + 1)] = np.mean(ramp_flow_collection_array[:, i])

                    # ------ get info. in the last interval
                    self.previous_ds_occ = np.mean(downstream_occ_collection_array[:, i])
                    self.previous_ramp_volume = np.mean(ramp_flow_collection_array[:, i])

                    # ------ predict the flow state in the next interval
                    # TODO for CXJ: we need prediction on downstream occ for each ramp
                    predicted_ds_occ = self.predict_flow()

                    # ------start of the controller------ #
                    # version 1----ALINEA algorithm
                    ramp_volume, green_time = self.cal_alinea_result()

                    # # version 2----ALINEA algorithm + Prediction
                    # ramp_volume, green_time = self.cal_alinea_result(prediction=predicted_ds_occ)
                    # ------end of the controller------ #

                    # save info. form controller
                    self.rm_data.loc[self.control_label, 'cal_ramp_volume' + str(i + 1)] = ramp_volume
                    self.rm_data.loc[self.control_label, 'green_time' + str(i + 1)] = green_time
                    self.amber_time[i] = self.time_interval_id * self.cycle_length + green_time
                    self.red_time[i] = self.time_interval_id * self.cycle_length + green_time + 3

                # at the start of every cycle, set the signal to green
                self.set_signal(target_key_list='all', state_list=['GREEN'])

            else:
                if self.current_time in self.amber_time:
                    key_list = []
                    for i, item in self.amber_time:
                        if item == self.current_time:
                            key_list.append(i)
                    self.set_signal(target_key_list=key_list, state_list=['AMBER' * len(key_list)])

                elif self.current_time in self.red_time:
                    key_list = []
                    for i, item in self.red_time:
                        if item == self.current_time:
                            key_list.append(i)
                    self.set_signal(target_key_list=key_list, state_list=['RED' * len(key_list)])

            # again keep the signal green
            self.set_signal(target_key_list='all', state_list=['GREEN'])

            # restore vehicle position update
        self.vissim_interface.ResumeUpdateGUI()
        # run single step to make uncontrolled vehicle to move
        self.vissim_interface.Simulation.RunSingleStep()

    def collect_scenario_data(self):
        """
        collection of basic static vissim data
        """
        self.get_route_links_set()

        # link id as key, link length as value
        for link in self.vissim_interface.Net.Links:
            self.link_length_map[str(link.AttValue("No"))] = float(link.AttValue("Length2D"))

    def init_vissim(self):
        # TODO for WSH: Vissim related settings

        # Load a Vissim Network:
        self.vissim_interface.LoadNet(r'.\vissim_network\xxx.inpx')

        # overall basic settings
        random_seed = 42
        self.vissim_interface.Simulation.SetAttValue('RandSeed', random_seed)
        self.vissim_interface.Simulation.SetAttValue('SimPeriod', self.end_of_simulation)
        self.vissim_interface.Simulation.SetAttValue('SimBreakAt', self.sim_break_at)
        # Set maximum speed
        self.vissim_interface.Simulation.SetAttValue('UseMaxSimSpeed', True)
        # Set sim_resolution
        self.vissim_interface.Simulation.SetAttValue('SimRes', 1 / self.delta_t)

        # settings for ramp and mainline
        quick_mode_label = True
        ramp_volume = 0
        mainline_volume = 0
        cc1_value = 1.35
        safety_reduction_factor = 0.35
        # self.vissim_interface.Net.VehicleInputs.ItemByKey(1).SetAttValue('Volume(1)', ramp_volume)
        # self.vissim_interface.Net.VehicleInputs.ItemByKey(2).SetAttValue('Volume(1)', mainline_volume)
        self.vissim_interface.Graphics.CurrentNetworkWindow.SetAttValue('QuickMode', quick_mode_label)
        # self.vissim_interface.Net.DrivingBehaviors.ItemByKey(3).SetAttValue('w99cc1', cc1_value)
        # self.vissim_interface.Net.DrivingBehaviors.ItemByKey(3).SetAttValue(
        #     'SafDistFactLnChg', safety_reduction_factor)

        # vissim data collection configuration, note that these parameters should be set before simulation starts
        self.vissim_interface.Evaluation.SetAttValue('DataCollCollectData', True)
        self.vissim_interface.Evaluation.SetAttValue('DataCollFromTime', 0)
        self.vissim_interface.Evaluation.SetAttValue('DataCollInterval', self.data_coll_interval)

    def get_all_veh_att(self):
        """
        get vehicle attribute from vissim
        :return:
        """

        self.vissim_interface.SuspendUpdateGUI()

        temp_all_veh_att = self.vissim_interface.Net.Vehicles.GetMultipleAttributes((
            "No", "SimSec", "VehType", "RoutDecNo", "RouteNo",
            "DesLane", "DestLane", "LeadTargNo", "FollowDist",
            "Hdwy", "LnChg", "Acceleration", "Speed", "Pos",
            "Lane", "PosLat"))

        write_label = 0
        for veh_att in temp_all_veh_att:
            if veh_att[3] is None:
                continue
            else:
                self.all_veh_attributes[write_label, 0] = veh_att[0]  # No
                self.all_veh_attributes[write_label, 1] = veh_att[1]  # SimSec
                self.all_veh_attributes[write_label, 2] = int(veh_att[2])  # VehType
                self.all_veh_attributes[write_label, 3] = veh_att[3]  # RoutDecNo
                route_dec_no = int(self.all_veh_attributes[write_label, 3])
                self.all_veh_attributes[write_label, 4] = veh_att[4]  # RouteNo
                route_no = int(self.all_veh_attributes[write_label, 4])
                self.all_veh_attributes[write_label, 5] = veh_att[5]  # DesLane
                self.all_veh_attributes[write_label, 6] = veh_att[6]  # DestLane
                self.all_veh_attributes[write_label, 7] = veh_att[7]  # LeadTargNo
                self.all_veh_attributes[write_label, 8] = veh_att[8]  # FollowDist
                self.all_veh_attributes[write_label, 9] = veh_att[9]  # Hdwy
                if veh_att[10] == 'NONE':
                    self.all_veh_attributes[write_label, 10] = 1  # LnChg
                elif veh_att[10] == 'LEFT':
                    self.all_veh_attributes[write_label, 10] = 2  # LnChg
                else:
                    self.all_veh_attributes[write_label, 10] = 3  # LnChg
                self.all_veh_attributes[write_label, 11] = veh_att[11]  # Acceleration
                self.all_veh_attributes[write_label, 12] = veh_att[12] / 3.6  # Speed (convert to m/s)
                self.all_veh_attributes[write_label, 13] = veh_att[13]  # Pos
                link_lane = [x for x in veh_att[14].split('-')]
                link_id = int(link_lane[0])
                self.all_veh_attributes[write_label, 14] = link_id  # link id
                lane_id = int(link_lane[1])
                self.all_veh_attributes[write_label, 15] = lane_id  # lane id
                self.all_veh_attributes[write_label, 16] = self.virtual_lane_id_map[link_id, lane_id]  # virtual lane id
                global_pos = self.link_pos_to_global_pos(route_dec_no, route_no, link_id, veh_att[13])
                self.all_veh_attributes[write_label, 17] = global_pos  # global_pos

            write_label += 1
        delete_label = range(write_label, 10001)
        self.all_veh_attributes = np.delete(self.all_veh_attributes, delete_label, 0)
        self.vissim_interface.ResumeUpdateGUI()

    def set_signal(self, target_key_list, state_list: list = None):
        """
        set the state of signal by key
        :param target_key_list: keys of signals to be changed
        :param state_list: target states of signals to be changed
        :return:
        """
        if target_key_list == 'all':
            # if get 'all' as the target_key,  set all signals as the same
            # in this case, state_list has only one state('GREEN', 'AMBER', 'RED')
            pass  # TODO
        else:
            # set signals in target_key_list according to the state_list
            for num, key in enumerate(target_key_list):
                # TODO: where to put 'num' ?
                self.vissim_interface.Net.SignalControllers.ItemByKey(1).SGs.ItemByKey(num). \
                    SetAttValue('State', state_list[num])

    def link_pos_to_global_pos(self, route_dec_in, route_in, link_in, link_pos_in):
        # the origin of global position is the start pos of first link in this route, different for each route
        global_pos = 0
        route_dec_and_route_id = str(route_dec_in) + str(route_in)
        route_links = self.route_links_set[route_dec_and_route_id]
        # note that the 0 position is at the merging point
        start_index = 0
        current_index = route_links.index(link_in)
        for i in range(start_index, current_index):
            global_pos += self.link_length_map[str(route_links[i])]
        global_pos += link_pos_in
        global_pos = global_pos - self.zeros_pos_route_coord[route_dec_and_route_id]
        return global_pos

    def get_loop_state(self) -> [list, list]:
        """
        TODO for WSH: get flow state from vissim
        :return: list of downstream occ and ramp flow
        """
        # to calculate downstream loops' average occupancy rate
        average_occ = []
        ramp_flow = []
        for i in range(self.ramp_num):
            # TODO: iterator represents the id of ramp. check where to put ’i‘ in the following iteration
            occ_result = [0, 0, 0, 0]
            lane_num = 3
            for lane_id in range(lane_num):
                data_col = self.vissim_interface.Net.DataCollectionMeasurements.ItemByKey(lane_id + 1)

                current_flow = data_col.AttValue('Vehs(Current,' + str(self.time_interval_id) + ',All)') * \
                               3600 / self.data_coll_interval
                current_speed = data_col.AttValue('Speed(Current,' + str(self.time_interval_id) + ',All)')
                current_veh_len = data_col.AttValue('Length(Current,' + str(self.time_interval_id) + ',All)')
                occ_result[lane_id] = current_flow / current_speed * (current_veh_len / 1000)
            average_occ[i] = sum(occ_result) / len(occ_result)
            ramp_flow[i] = self.vissim_interface.Net.DataCollectionMeasurements.ItemByKey(5).AttValue(
                'Vehs(Current,' + str(self.time_interval_id) + ',All)') * 3600 / self.data_coll_interval
        return average_occ, ramp_flow

    def cal_alinea_result(self, prediction=None):
        # calculate the green time for ramp signal using ALINEA algorithm
        # TODO for ZXC: add funcs for using prediction results
        kr = 70
        critical_occ = 0.20
        r_min = 500
        r_max = self.saturated_flow
        new_r = self.previous_ramp_volume + kr * (critical_occ - self.previous_ds_occ)
        new_r = max(r_min, new_r)
        new_r = min(r_max, new_r)
        # what if new_r is negative???
        # then we should set a minimum green time (e.g. 10)
        current_green_time = max(10, round(new_r / self.saturated_flow * self.cycle_length))
        current_green_time = min(current_green_time, self.cycle_length - 3)
        return [new_r, current_green_time]

    def predict_flow(self) -> list:
        """
        TODO for CXJ
        TODO: check get_loop_state() for possibly useful information from vissim
        predict occ of mainline in the next control interval using current flow state
        :return: [list] of occ at five ramp's downstream
        """
        prev_occ = self.previous_ds_occ
        current_occ = prev_occ

        return current_occ

    def get_route_links_set(self):
        # output is a dictionary, whose keys are str(combine route_dec_id and route_id) and
        # values are a list of sequential links id
        output_dic = {}
        for route_dec in self.vissim_interface.Net.VehicleRoutingDecisionsStatic:
            routes_container = route_dec.VehRoutSta
            start_link_id = [int(route_dec.AttValue("Link"))]  # convert to int then put it in a list
            for single_route in routes_container:
                route_dec_id = single_route.AttValue("VehRoutDec")  # the return value is a unicode str
                route_id = single_route.AttValue("No")  # the return value is a unicode str
                links_seq = [int(x) for x in single_route.AttValue("LinkSeq").split(',')]
                end_link_id = [int(single_route.AttValue("DestLink"))]
                output_dic[str(route_dec_id) + str(route_id)] = start_link_id + links_seq + end_link_id
        self.route_links_set = output_dic


if __name__ == '__main__':

    simu = Simulator()
    simu.simulate_scenario()
