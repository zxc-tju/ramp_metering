import win32com.client as com
import numpy as np
import pandas as pd
import datetime
import xlwt
import openpyxl
import os


class Simulator:

    def __init__(self):

        # ------ static simulation settings
        self.number_of_step = 4000  # 作用是啥？
        self.sim_break_at = 500  # 需要修改吗？
        self.end_time_of_simulation = 5000
        self.end_time_of_control = 4400
        self.delta_t = 1  # simulation resolution
        self.data_coll_interval = 40

        # ------ simulator initialize
        self.current_time = self.sim_break_at
        self.vissim_interface = None
        self.all_time_all_veh_trj = None
        self.previous_ds_occ = None
        self.previous_ramp_volume = None
        self.control_label = 0
        self.time_interval_id = 0
        self.global_trj_write_label = 0
        self.total_trj_record_num = 10000000
        self.link_length_map = {}
        self.route_links_set = {}
        self.amber_time = []
        self.red_time = []
        self.downstream_flow_collection = []
        self.downstream_speed_collection = []
        self.downstream_occ_collection = []
        self.ramp_flow_collection = []
        self.ramp_speed_collection = []

        headings_list = [['downstream_occ-' + str(i + 1),
                          'ramp_flow-' + str(i + 1),
                          'cal_ramp_volume-' + str(i + 1),
                          'green_time-' + str(i + 1)] for i in range(5)]
        headings = np.array(headings_list).reshape([-1])
        # for save data in each control interval
        self.rm_data = pd.DataFrame(np.zeros((self.number_of_step, 20)), columns=headings)
        # for save vehicle trajectories
        # self.all_veh_attributes = []
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
        # TODO for WSH: check following info. for the competition scenario   没看懂 最后解决
        self.ramp_num = 5  # number of ramps
        self.saturated_flow = 1800  # single lane saturated flow
        self.cycle_length = 40
        # self.virtual_lane_id_map = {(1, 1): 1, (10000, 1): 1, (2, 1): 1, (10001, 1): 1,
        #                             (3, 1): 1, (3, 2): 1, (10003, 1): 1,
        #                             (10002, 1): 1, (4, 1): 1, (2, 2): 2, (10001, 2): 2,
        #                             (3, 3): 2, (2, 2): 2, (10002, 2): 2,
        #                             (4, 2): 2, (2, 3): 3, (2, 3): 3, (10001, 3): 3,
        #                             (3, 4): 3, (10002, 3): 3, (4, 3): 3,
        #                             (2, 4): 4, (10001, 4): 4, (3, 5): 4,
        #                             (10002, 4): 4, (4, 4): 4}
        # self.zeros_pos_route_coord = {'11': 268.35, '21': 339.458}
        # self.zeros_pos_route_coord = {}

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
            self.current_time += self.delta_t

            # ------ get and save trajectories at each step
            # all_veh_attributes_list = self.get_all_veh_att()
            # if self.global_trj_write_label <= self.total_trj_record_num - 1:
            #     start_label = self.global_trj_write_label
            #     self.all_time_all_veh_trj[start_label:start_label + len(all_veh_attributes_list[:, 0]), :] \
            #         = all_veh_attributes_list
            #     self.global_trj_write_label += len(all_veh_attributes_list[:, 0])   #  这一段的用处是啥？？？

            # ------ get and save flow state for each data_coll_interval (4 seconds)
            if self.current_time % self.data_coll_interval == 0:

                # TODO for WSH: we need flow info. of all five ramps using 'get loop state()'  DONE
                # ------ get time interval id
                self.time_interval_id = int(self.current_time / self.cycle_length)
                downstream_flow_list, downstream_speed_list, downstream_occ_list, ramp_flow_list, ramp_speed_list = self.get_loop_state()

                # store new data into the collection
                self.downstream_flow_collection.append(downstream_flow_list)
                self.downstream_speed_collection.append(downstream_speed_list)
                self.downstream_occ_collection.append(downstream_occ_list)
                self.ramp_flow_collection.append(ramp_flow_list)
                self.ramp_speed_collection.append(ramp_speed_list)

                # delete old data, keep only 10 interval in the collection
                if len(self.downstream_occ_collection) == 10:
                    del self.downstream_occ_collection[0]
                    del self.ramp_flow_collection[0]

            "Beginning of a control interval"
            if self.current_time % self.cycle_length == 0:

                # ------ update control interval id
                self.control_label = self.control_label + 1

                # # ------ get time interval id
                # self.time_interval_id = int(self.current_time / self.cycle_length)

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
                    self.amber_time.append(self.time_interval_id * self.cycle_length + green_time)     # time_interval_i是第几个周期，可以乘吗？
                    self.red_time.append(self.time_interval_id * self.cycle_length + green_time + 3)

                # at the start of every cycle, set the signal to green
                self.set_signal(target_key_list='all', state_list=['GREEN'])

            else:
                if self.current_time in self.amber_time:
                    key_list = []
                    for i, item in enumerate(self.amber_time):
                        if item == self.current_time:
                            key_list.append(i)
                    self.set_signal(target_key_list=key_list, state_list=['AMBER'])

                elif self.current_time in self.red_time:
                    key_list = []
                    for i, item in enumerate(self.red_time):
                        if item == self.current_time:
                            key_list.append(i)
                    self.set_signal(target_key_list=key_list, state_list=['RED'])

            # restore vehicle position update
            self.vissim_interface.ResumeUpdateGUI()
            # run single step to make uncontrolled vehicle to move
            self.vissim_interface.Simulation.RunSingleStep()

            if self.current_time == self.end_time_of_control:
                # again keep the signal green
                self.set_signal(target_key_list='all', state_list=['GREEN'])

        # end of simulation
        self.vissim_interface.Simulation.Stop()

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
        self.vissim_interface.LoadNet(r'C:\Users\Administrator\Desktop\VISSIM_tram\仿真文件\WSH\华为比赛\Vissim\freeway.inpx')

        # overall basic settings
        random_seed = 42
        self.vissim_interface.Simulation.SetAttValue('RandSeed', random_seed)
        self.vissim_interface.Simulation.SetAttValue('SimPeriod', self.end_time_of_simulation)
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

    def get_all_veh_att(self)-> [list]:
        """
        get vehicle attribute from vissim
        :return:
        """

        self.vissim_interface.SuspendUpdateGUI()
        all_veh_attributes = []
        Record_veh_trajectory = []

        temp_all_veh_att = self.vissim_interface.Net.Vehicles.GetMultipleAttributes((
            "No", "SimSec", "VehType", "RoutDecNo", "RouteNo",
            "DesLane", "DestLane", "LeadTargNo", "FollowDist",
            "Acceleration", "Speed", "Pos",
            "Lane", "PosLat"))

        #write_label = 0
        for veh_att in temp_all_veh_att:
            one_veh = []
            if veh_att[3] is None:
                continue     # 不考虑不到路径决策点的车。
            else:
                one_veh.append(veh_att[0])  # No
                one_veh.append(veh_att[1])  # SimSec
                one_veh.append(int(veh_att[2]))  # VehType
                one_veh.append(veh_att[3])  # RoutDecNo
                route_dec_no = int(one_veh[3])
                one_veh.append(veh_att[4])  # RouteNo
                route_no = int(one_veh[4])
                one_veh.append(veh_att[5])  # DesLane
                one_veh.append(veh_att[6]) # DestLane
                one_veh.append(veh_att[7])  # LeadTargNo
                one_veh.append(veh_att[8])  # FollowDist

                one_veh.append(veh_att[9])  # Acceleration
                one_veh.append(veh_att[10] / 3.6)  # Speed (convert to m/s)
                one_veh.append(veh_att[11])  # Pos
                link_lane = [x for x in veh_att[12].split('-')]
                link_id = int(link_lane[0])
                one_veh.append(link_id)  # link id
                lane_id = int(link_lane[1])
                one_veh.append(lane_id)  # lane id
                one_veh.append(link_lane)  # lane id
                # all_veh_attributes[write_label][15] = self.virtual_lane_id_map[link_id, lane_id]  # virtual lane id
                # global_pos = self.link_pos_to_global_pos(route_dec_no, route_no, link_id, veh_att[11])
                # all_veh_attributes[write_label][15] = global_pos  # global_pos

                all_veh_attributes.append(one_veh)

            #write_label += 1
        #delete_label = range(write_label, 10001)
        #self.all_veh_attributes = np.delete(self.all_veh_attributes, delete_label, 0)

        # save vehicle information
        cur_t = datetime.datetime.now()
        minor_path = 'All_veh_tra' + str(cur_t.month) + '-' + str(cur_t.day) + '-' + str(cur_t.hour) + '-' + str(
            cur_t.minute) + '.txt'
        # 创建txt
        file = open('./veh_data/' + minor_path, 'w')
        file.write(
            '"No", "SimSec", "VehType", "RoutDecNo", "RouteNo", "DesLane", "DestLane", "LeadTargNo", "FollowDist", '
            '"Acceleration", "Speed", "Pos", "linkID", "laneID", "linklaneID"' + os.linesep)

        for i in all_veh_attributes:
            Record_veh_trajectory.append(i)
        for ii in Record_veh_trajectory:
            file.write(str(ii)[1:-1] + '\n')
        file.close()

        print('所有小汽车轨迹已保存至：' + minor_path)

        self.vissim_interface.ResumeUpdateGUI()
        return all_veh_attributes

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
            for num in range(self.ramp_num):
                # TODO: where to put 'num' ?
                self.vissim_interface.Net.SignalControllers.ItemByKey(num+1).SGs.ItemByKey(1). \
                    SetAttValue('State', state_list[0])
            pass  # TODO
        else:
            # set signals in target_key_list according to the state_list
            for num, key in enumerate(target_key_list):
                # TODO: where to put 'num' ?
                self.vissim_interface.Net.SignalControllers.ItemByKey(num+1).SGs.ItemByKey(1). \
                    SetAttValue('State', state_list[0])

    # def link_pos_to_global_pos(self, route_dec_in, route_in, link_in, link_pos_in):
    #     # the origin of global position is the start pos of first link in this route, different for each route
    #     global_pos = 0
    #     route_dec_and_route_id = str(route_dec_in) + str(route_in)
    #     route_links = self.route_links_set[route_dec_and_route_id]
    #     # note that the 0 position is at the merging point
    #     start_index = 0
    #     current_index = route_links.index(link_in)
    #     for i in range(start_index, current_index):
    #         global_pos += self.link_length_map[str(route_links[i])]
    #     global_pos += link_pos_in
    #     global_pos = global_pos - self.zeros_pos_route_coord[route_dec_and_route_id]
    #     return global_pos

    def get_loop_state(self) -> [list, list, list, list, list]:
        """
        # TODO for WSH: get flow state from vissim    # time_interval_id 有些迷惑
        :return: list of downstream flow, speed and ramp flow, speed
        """
        # to calculate downstream loops' average occupancy rate
        downstream_flow = []
        downstream_speed =[]
        downstream_occ = []
        ramp_flow = []
        ramp_speed = []
        # downstream flow speed output
        for i in range(2,7):
            data_col = self.vissim_interface.Net.DataCollectionMeasurements.ItemByKey(i)

            current_flow = data_col.AttValue('Vehs(Current,' + str(self.time_interval_id) + ',All)')\
                           * 3600 / self.data_coll_interval
            current_speed =  data_col.AttValue('Speed(Current,' + str(self.time_interval_id) + ',All)')
            current_veh_len = data_col.AttValue('Length(Current,' + str(self.time_interval_id) + ',All)')

           # in case no vehicle pass
            if current_flow == 0:
                current_speed = 0
                current_downstream_occ = 0
            else:
                current_downstream_occ = current_flow / current_speed * (current_veh_len / 1000)

            downstream_flow.append(current_flow)
            downstream_speed.append(current_speed)
            downstream_occ.append(current_downstream_occ)

        # ramp flow speed output
        for j in [8,10,12,14,16]:
            data_col_ramp = self.vissim_interface.Net.DataCollectionMeasurements.ItemByKey(j)
            ramp_flow_current = data_col_ramp.AttValue('Vehs(Current,' + str(self.time_interval_id) + ',All)') * \
                                   3600 / self.data_coll_interval

            # in case no vehicle pass
            if ramp_flow_current == 0:
                ramp_speed_current = 0
            else:
                ramp_speed_current =data_col_ramp.AttValue('Speed(Current,' + str(self.time_interval_id) + ',All)')

            ramp_flow.append(ramp_flow_current)
            ramp_speed.append(ramp_speed_current)

        return downstream_flow, downstream_speed, downstream_occ, ramp_flow, ramp_speed

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
