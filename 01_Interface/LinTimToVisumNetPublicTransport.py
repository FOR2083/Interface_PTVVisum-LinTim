# -*- coding: iso-8859-1 -*-
''' TP1_NET_PUBLIC_TRANSPORT
INPUT:          class ReadLinTimFiles, class VisumFolderStructure
OUTPUT:         public_transport.net
DESCRIPTION:    creates a net file for the public transport information like lines, line routes, line route items,
                time profiles, time profile items, vehicle journeys, vehicle journey sections
                if script is executed in the procedure sequence it imports the net-file to Visum
                fixed-lines are ignored - use script 'import_fixed_lines_lintim.py'
'''

import os
import time
import pandas as pd

# ========== classes =========
class VisumTableHeader:
    vision_en = "$VISION\n" \
                "* Universität Stuttgart Fakultät 2 Bau+Umweltingenieurwissenschaften Stuttgart\n" \
                "* 04/25/16\n"
    table_version_block_net_en = "*\n" \
                                 "* Table: Version block\n" \
                                 "*\n" \
                                 "$VERSION:VERSNR;FILETYPE;LANGUAGE;UNIT\n" \
                                 "10,000;Net;ENG;KM\n"
    table_lines_en = "*\n" \
                     "* Table: Lines\n" \
                     "*\n" \
                     "$LINE:NAME;TSYSCODE\n"
    table_line_routes_en = "*\n" \
                           "* Table: Line routes\n" \
                           "*\n" \
                           "$LINEROUTE:LINENAME;NAME;DIRECTIONCODE;ISCIRCLELINE\n"
    table_line_route_items_en = "*\n" \
                                "* Table: Line route items\n" \
                                "*\n" \
                                "$LINEROUTEITEM:LINENAME;LINEROUTENAME;DIRECTIONCODE;INDEX;ISROUTEPOINT;NODENO;STOPPOINTNO\n"
    table_time_profiles_en = "*\n" \
                             "* Table: Time profiles\n" \
                             "*\n" \
                             "$TIMEPROFILE:LINENAME;LINEROUTENAME;DIRECTIONCODE;NAME\n"
    table_time_profile_items_en = "*\n" \
                                  "* Table: Time profile items\n" \
                                  "*\n" \
                                  "$TIMEPROFILEITEM:LINENAME;LINEROUTENAME;DIRECTIONCODE;TIMEPROFILENAME;INDEX;LRITEMINDEX;ALIGHT;BOARD;ARR;DEP\n"
    table_vehicle_journeys_en = "*\n" \
                                "* Table: Vehicle journeys\n" \
                                "*\n" \
                                "$VEHJOURNEY:NO;DEP;LINENAME;LINEROUTENAME;DIRECTIONCODE;TIMEPROFILENAME;FROMTPROFITEMINDEX;TOTPROFITEMINDEX\n"
    table_vehicle_journey_items_en = "*\n" \
                                     "* Table: Vehicle journey items\n" \
                                     "*\n" \
                                     "$VEHJOURNEYITEM:VEHJOURNEYNO;INDEX\n"
    table_vehicle_journey_sections_en = "*\n" \
                                        "* Table: Vehicle journey sections\n" \
                                        "*\n" \
                                        "$VEHJOURNEYSECTION:VEHJOURNEYNO;NO;FROMTPROFITEMINDEX;TOTPROFITEMINDEX\n"
    table_network_lintim_en = "*\n" \
                    "* Table: Network\n" \
                    "*\n" \
                    "$NETWORK:LINTIM_OVERALL_RUNTIME\n"

class VisumFolderStructure:
    def __init__(self, Visum, sep_folder):
        self.main = Visum.UserPreferences.DocumentName.rsplit(sep_folder, 2)[0]
        self.version = Visum.GetPath(2)
        self.network = Visum.GetPath(1)
        self.matrix = Visum.GetPath(69)
        self.graphicparameter = Visum.GetPath(8)
        self.lt_input = sep_folder.join([self.main, '_interface', "LinTim2Visum"])
        self.lt_version_LinTim = sep_folder.join([self.lt_input, Visum.Net.AttValue("lintim_version_name")])
        self.lt_version_basis = sep_folder.join([self.lt_version_LinTim, 'basis'])
        self.lt_version_timetabling = sep_folder.join([self.lt_version_LinTim, 'timetabling'])
        self.lt_version_lineplanning = sep_folder.join([self.lt_version_LinTim, 'line-planning'])

class ReadLinTimFiles:
    ''' reads all lintim files
        add dictionary for id translation from stop.giv file
        add dictionary for load
        add hierarchical line representation
    '''

    def __init__(self, folder_paths, lintim_version, Visum):
        value_transport_systems_code = Visum.Net.AttValue("lintim_tsys_for_adapting")
        self.folder_paths = folder_paths
        self.lintim_version = lintim_version
        self.lt_stops = open(os.path.join(self.folder_paths.lt_version_basis,
                                          ValueLinTimFile.stop
                                          ), "r").readlines()
        self.lt_stops_dic_id = {int(item.rsplit(';')[0]): int(item.rsplit(';')[1]) for item in self.lt_stops if
                                not item.startswith("#")}
        self.lt_timetable = open(os.path.join(self.folder_paths.lt_version_timetabling,
                                              ValueLinTimFile.timetable
                                              ), "r").readlines()

        # Import timetable
        self.lt_line = TimetableReader(filename=os.path.join(self.folder_paths.lt_version_timetabling,
                                                             ValueLinTimFile.timetable
                                                             ), Visum=Visum)
        # formatting
        self.lt_line = self.lt_line.data_reformed

        # Read fixed lines
        try:
            # read file "line-Names.lin" as lintim output
            df_fixed_lines = pd.read_csv(os.path.join(self.folder_paths.lt_version_lineplanning,
                                                                 ValueLinTimFile.fixed_lines)
                                           , delimiter=";")

            # remove # & " " in column names
            df_fixed_lines.columns = df_fixed_lines.columns.astype(str).str.replace("[ ]", "").str.replace("[#]", "")
            df_fixed_lines.loc[:, "name"] = df_fixed_lines["name"].str.lstrip()
            df_fixed_lines[["linename", "route", "timeprofiles"]] = df_fixed_lines.loc[:, "name"].str.split("-",
                                                                                                            expand=True)
            df_fixed_lines[["tp1", "tp2"]] = df_fixed_lines.loc[:, "timeprofiles"].str.split("_", expand=True)
            df_fixed_lines.index = df_fixed_lines["line-id"]
            # save ids of fixed lines in set
            self.fixed_lines = set(df_fixed_lines.loc[:, "line-id"].unique())

            # chance line id/name to original (visum) names
            self.lt_line.update({df_fixed_lines.loc[line, "linename"]: values for line, values in self.lt_line.items() if line in self.fixed_lines})
            # remove old ids from dict
            self.lt_line = {line: values for line, values in self.lt_line.items() if line not in self.fixed_lines}

        except:
            # if file is not existing: imitialize empty set
            self.fixed_lines = set()

        # add tsys for lines
        self.line_tsys = {line: value_transport_systems_code for line in self.lt_line.keys() if line not in self.fixed_lines}

        # add tsys for fixed-lines
        # import tsys from visum
        attr_lines = ["Name", "TSysCode"]
        df_fixed_lines_visum = pd.DataFrame(Visum.Net.Lines.GetMultipleAttributes(attr_lines), columns=attr_lines)
        self.line_tsys.update({line["Name"]: line["TSysCode"] for id, line in df_fixed_lines_visum.iterrows()})

        # read lintim values if existing
        try:
            self.lt_values = open(os.path.join(self.folder_paths.lt_version_LinTim,
                                              ValueLinTimFile.values
                                              ), "r").readlines()

            self.lt_values_dic_id = {item.rsplit(':')[0]: item.split(': ')[-1].replace('\n', '') for item in self.lt_values if not item.startswith("#")}
        except:
            self.lt_values_dic_id = {}

class ValueLinTimFile:
    stop = "Stop.giv"
    edge = "Edge.giv"
    load = "Load.giv"
    od = "OD.giv"
    fixed_lines = "Line-Names.lin"
    timetable = "Timetable-visum-nodes.tim"
    values = "values.giv"

class TimetableReader:
    '''
        converts the list representation of the timetable into a hierarchical tree representation of lines
        checks the format version, either with the frequency(old) or the exact departure times(new)
        polish the raw data to directly use it for net file writing
    '''

    def __init__(self, filename, Visum):
        '''
            read data into pd dataset
            check which version of line representation is given
        '''
        self.df_raw = pd.read_csv(filename, sep=';', skipinitialspace=True)
        self.col_group_attr = ['line-id', 'line-code', 'direction', 'frequency', 'line-repetition']

        # get units per minute
        self.time_units_per_minute = Visum.Net.AttValue("lintim_time_units_per_minute")
        # consider period length in minutes
        self.period_length_minutes = Visum.Net.AttValue("lintim_period_length") / self.time_units_per_minute

        # format column names
        self.df_raw.columns = self.df_raw.columns.str.replace('#', '')
        self.df_raw.columns = self.df_raw.columns.str.strip()  # remove whitespace from column header

        # translate data to Visum PuT supply elements
        self.data_reformed = self.data_polish()
        self.data_polish_timeprofile(Visum)

        if self.time_units_per_minute != 60:
            # convert time unit representation to seconds
            self.data_polish_to_seconds()

        # expand timetable to assignment period
        self.data_polish_expand_timetable()

    def data_polish(self):
        '''
            takes the raw data and converts the list representation of the timetable into a more VISUM common tree line representation
        :return: dict of lines, lineroutes, ...
        '''

        '''
            group data by line-code+line-repetition
            set multi index for line-id+line-code+line-repetition
            convert dataframe into nested dict
        '''
        df_result = self.df_raw.groupby(['line-code', 'line-repetition']).agg({'line-id': 'first',
                                                                               'line-code': 'first',
                                                                               'line-repetition': 'first',
                                                                               "direction": 'first',
                                                                               "stop-order": list,
                                                                               "stop-id": list,
                                                                               "frequency": 'first',
                                                                               "arrival_time": list,
                                                                               "departure_time": list})
        df_result.set_index(['line-id', 'line-code', 'line-repetition'], inplace=True)  # multiple index

        # transform dataframe to dictionary in 4 steps
        # 1. dataframe to nested dict (multi indices as keys)
        # 2. add information about general time profile on line route dictionary
        # 3. add attribute "direction" and "frequency" on line route dictionary
        # 4. delete inner dict for trips

        # 1. dataframe to nested dict (multi indices as keys)
        dict_result = {}
        for line_name in df_result.index.levels[0]:
            if line_name not in dict_result:
                dict_result[line_name] = {}
            for line_route in df_result.xs(line_name).reset_index()["line-code"]:
                if line_route not in dict_result[line_name]:
                    dict_result[line_name][line_route] = {}
                for no_trip_in_period in df_result.xs(line_name).xs(line_route).reset_index()["line-repetition"]:
                    if no_trip_in_period not in dict_result[line_name][line_route]:
                        dict_result[line_name][line_route][no_trip_in_period] = {}
                        dict_result[line_name][line_route][no_trip_in_period] = df_result.xs(line_name).xs(line_route).xs(no_trip_in_period).to_dict()

        # 2. add information about general time profile on line route dictionary
        for key_line in dict_result:
            for key_lineroute in dict_result[key_line]:
                trips = []
                for key_trip in dict_result[key_line][key_lineroute]:
                    trips.append(dict_result[key_line]
                                 [key_lineroute]
                                 [key_trip]
                                 ['arrival_time']
                                 [0])  # get first arrival time for each line repetition and convert minutes to seconds
                trips.sort()
                dict_result[key_line][key_lineroute].update(
                    {'time_profile_trips_hour': trips})  # trips: get first arrivial time for each line-repetition
                dict_result[key_line][key_lineroute].update(
                    {'time_profile_arr': dict_result[key_line][key_lineroute][1]['arrival_time']})
                dict_result[key_line][key_lineroute].update(
                    {'time_profile_dep': dict_result[key_line][key_lineroute][1]['departure_time']})
                dict_result[key_line][key_lineroute].update(
                    {'time_profile_stop_order': dict_result[key_line][key_lineroute][1]['stop-order']})
                dict_result[key_line][key_lineroute].update(
                    {'time_profile_stop_id': dict_result[key_line][key_lineroute][1]['stop-id']})


        # 3. add attribute "direction" and "frequency" on line route dictionary
        # 4. delete inner dict for trips
        for key_line in dict_result:
            for key_lineroute in dict_result[key_line]:
                # frequency
                dict_result[key_line][key_lineroute].update({'frequency':
                                                                 dict_result[key_line]
                                                                 [key_lineroute]
                                                                 [1]
                                                                 ['frequency']})
                # direction
                dict_result[key_line][key_lineroute].update({'direction':
                                                                 dict_result[key_line]
                                                                 [key_lineroute]
                                                                 [1]
                                                                 ['direction']})

                # delete dictionary of trip numbers
                for i in range(1, len(dict_result[key_line][key_lineroute])):
                    # delete if key exists, else do nothing
                    dict_result[key_line][key_lineroute].pop(i, None)

        return dict_result

    def data_polish_timeprofile(self, Visum):
        '''
            - define the first arr/dep time to reduce the timeprofile by this value
            - old: 21,30,36,... new: 0,9,15,... the clean time profile is the input for the Visum time profile generation
        :return:
        '''
        for key_line in self.data_reformed:
            for key_lineroute in self.data_reformed[key_line]:
                # arrival
                # time of the departure time at the first stop, it is needed to start the time profile item at minute = 0
                self.data_reformed[key_line][key_lineroute].update({'time_profile_arr_first':
                                                                        self.data_reformed[key_line]
                                                                        [key_lineroute]
                                                                        ['time_profile_arr']
                                                                        [0]})

                # arrival time - first arrival time at first stop for all stops
                self.data_reformed[key_line][key_lineroute]['time_profile_arr'] = \
                    [arr - self.data_reformed[key_line][key_lineroute]['time_profile_arr_first'] for arr in
                     self.data_reformed[key_line][key_lineroute][
                         'time_profile_arr']]

                # adjust factor at period change, s.t. arrival[x] > arrival[x-1]
                factor = 0
                for idx, arr in enumerate(self.data_reformed[key_line][key_lineroute]['time_profile_arr']):
                    if not idx == 0:
                        if self.data_reformed[key_line][key_lineroute]['time_profile_arr'][idx - 1] > \
                                arr + self.period_length_minutes * self.time_units_per_minute * factor:
                            factor += 1
                        self.data_reformed[key_line][key_lineroute]['time_profile_arr'][idx] += \
                            self.period_length_minutes * self.time_units_per_minute * factor

                # departure
                self.data_reformed[key_line][key_lineroute].update(
                    {'time_profile_dep_first': self.data_reformed[key_line][key_lineroute]['time_profile_dep'][0]})
                self.data_reformed[key_line][key_lineroute]['time_profile_dep'] = [
                    dep - self.data_reformed[key_line][key_lineroute]['time_profile_dep_first'] for dep in
                    self.data_reformed[key_line][key_lineroute][
                        'time_profile_dep']]

                # adjust factor at period change, s.t. departurel[x] > departure[x-1]
                factor = 0
                for idx, dep in enumerate(self.data_reformed[key_line][key_lineroute]['time_profile_dep']):
                    if not idx == 0:
                        if self.data_reformed[key_line][key_lineroute]['time_profile_dep'][idx - 1] > \
                                dep + self.period_length_minutes * self.time_units_per_minute * factor:
                            factor += 1
                        self.data_reformed[key_line][key_lineroute]['time_profile_dep'][idx] += \
                            self.period_length_minutes * self.time_units_per_minute * factor


    def data_polish_to_seconds(self):
        factor_time_unit_to_seconds = 60 / self.time_units_per_minute

        for key_line in self.data_reformed:
            for key_lineroute in self.data_reformed[key_line]:
                self.data_reformed[key_line][key_lineroute]['time_profile_arr'] = [i * factor_time_unit_to_seconds for i in
                                                                                   self.data_reformed[key_line][
                                                                                       key_lineroute][
                                                                                       'time_profile_arr']]
                self.data_reformed[key_line][key_lineroute]['time_profile_arr_first'] *= factor_time_unit_to_seconds
                self.data_reformed[key_line][key_lineroute]['time_profile_dep'] = [i * factor_time_unit_to_seconds for i in
                                                                                   self.data_reformed[key_line][
                                                                                       key_lineroute][
                                                                                       'time_profile_dep']]
                self.data_reformed[key_line][key_lineroute]['time_profile_dep_first'] *= factor_time_unit_to_seconds
                self.data_reformed[key_line][key_lineroute]['time_profile_trips_hour'] = [i * factor_time_unit_to_seconds for i in
                                                                                          self.data_reformed[key_line][
                                                                                              key_lineroute][
                                                                                              'time_profile_trips_hour']]

    def data_polish_expand_timetable(self):
        for key_line in self.data_reformed:
            for key_lineroute in self.data_reformed[key_line]:
                self.data_reformed[key_line][key_lineroute]['time_profile_trips_period'] = \
                    [period + trip for trip in self.data_reformed[key_line][key_lineroute]['time_profile_trips_hour']
                     for period in range(value_schedule_creation_from,
                                         value_schedule_creation_to + 1,
                                         3600)]

# ======== functions =========

def import_net2visum(path_net):
    # create NetRouteSearchTSys-Object and choose route search options
    NetReadRouteSearchTSysController = Visum.CreateNetReadRouteSearchTSys()
    NetReadRouteSearchTSysController.SetAttValue("HOWTOHANDLEINCOMPLETEROUTE", "SEARCHSHORTESTPATH")
    NetReadRouteSearchTSysController.SetAttValue("SHORTESTPATHCRITERION", 2)
    NetReadRouteSearchTSysController.SetAttValue("INCLUDEBLOCKEDLINKS", False)
    NetReadRouteSearchTSysController.SetAttValue("INCLUDEBLOCKEDTURNS", False)
    NetReadRouteSearchTSysController.SetAttValue("MAXDEVIATIONFACTOR", 20)
    NetReadRouteSearchTSysController.SetAttValue("WHATTODOIFSTOPPOINTNOTFOUND", 1)
    NetReadRouteSearchTSysController.SetAttValue("LINKTYPEFORINSERTEDLINKSREPLACINGMISSINGSHORTESTPATHS", 98)

    # create NetRouteSearch-Object and assign NetRouteSearchTSys-objects
    NetReadRouteSearchController = Visum.CreateNetReadRouteSearch()
    NetReadRouteSearchController.SetForAllTSys(NetReadRouteSearchTSysController)

    # create AddNetRead-Object and specify desired conflict avoiding method
    AddNetReadController = Visum.CreateAddNetReadController
    AddNetReadController.SetWhatToDo("Network", 5)

    Visum.IO.LoadNet(NetFile=path_net,
                     ReadAdditive=True,
                     routeSearch=NetReadRouteSearchController,
                     AddNetRead=AddNetReadController)

# ======== CORE =========
''' VisumStart
- enables the usage in the procedure sequence or as Visum instance
- name of the Visum version that is opened
'''

try:
    global Visum
    Visum
except NameError:
    import win32com.client as com

    Visum = com.Dispatch("Visum.Visum.200")
    name = r'E:\Seafile\36_VRS\Version\debug_import.ver'
    Visum.LoadVersion(name)

## set attributes
# start of daily PuT supply
value_schedule_creation_from = 5 * 3600
# end of daily PuT supply
value_schedule_creation_to = 22 * 3600
# define separator of folder in folder pathes
sep_folder = "\\"

folder_paths = VisumFolderStructure(Visum, sep_folder)

# get attributes
scenario_name = Visum.Net.AttValue("scenario_name")
lintim_version = Visum.Net.AttValue("lintim_version_name")

# reads all LinTim files and stores the values in the ReadLinTimFiles class
lintim_files = ReadLinTimFiles(folder_paths, lintim_version, Visum)

name_put_supply_net = scenario_name + "_" + "Lines" + "_" + lintim_version + ".net"

# WRITE NET FILE FOR PUBLIC TRANSPORT SYSTEM
with open( sep_folder.join([folder_paths.network, name_put_supply_net]), "w") as f:
    # add header
    f.write(VisumTableHeader.vision_en)
    # Net Header
    f.write(VisumTableHeader.table_version_block_net_en)

    # lines
    f.write(VisumTableHeader.table_lines_en)
    for line_key, line_val in lintim_files.lt_line.items():
        f.write(';'.join([str(line_key),  # line name
                          lintim_files.line_tsys[line_key] + '\n']))  # tsys code

    # line routes
    f.write(VisumTableHeader.table_line_routes_en)
    for line_key, line_val in lintim_files.lt_line.items():
        for lineroute_key, lineroute_val in line_val.items():
            f.write(';'.join([str(line_key),  # line name
                              str(lineroute_key),  # line route name
                              lineroute_val['direction'],  # direction code
                              str(0) + '\n']))  # circle line

    # line route items
    f.write(VisumTableHeader.table_line_route_items_en)
    for line_key, line_val in lintim_files.lt_line.items():
        for lineroute_key, lineroute_val in line_val.items():
            for stop_index, stop_sequence in zip(lineroute_val['time_profile_stop_order'],
                                                 lineroute_val['time_profile_stop_id']):
                f.write(';'.join([str(line_key),  # line name
                                  str(lineroute_key),  # line route name
                                  lineroute_val['direction'],  # direction code
                                  str(stop_index),  # index of the stop
                                  str(1),  # is route stop, 1=yes
                                  str(lintim_files.lt_stops_dic_id[stop_sequence]),
                                  # node stop no converted to Visum stop no
                                  str(lintim_files.lt_stops_dic_id[
                                          stop_sequence]) + '\n']))  # stop point no converted to Visum stop no
    # time profiles
    f.write(VisumTableHeader.table_time_profiles_en)
    for line_key, line_val in lintim_files.lt_line.items():
        for lineroute_key, lineroute_val in line_val.items():
            f.write(';'.join([str(line_key),  # line name
                              str(lineroute_key),  # line route name
                              lineroute_val['direction'],  # direction code
                              '_'.join([str(line_key), str(lineroute_key)]) + '\n']))  # time profile name

    # time profile items
    f.write(VisumTableHeader.table_time_profile_items_en)
    for line_key, line_val in lintim_files.lt_line.items():
        for lineroute_key, lineroute_val in line_val.items():
            for stop_index, arr, dep in zip(lineroute_val['time_profile_stop_order'],
                                            lineroute_val['time_profile_arr'],
                                            lineroute_val['time_profile_dep']):

                f.write(';'.join([str(line_key),  # line name
                                  str(lineroute_key),  # line route name
                                  lineroute_val['direction'],  # direction code
                                  '_'.join([str(line_key), str(lineroute_key)]),  # time profile name
                                  str(stop_index),  # index of the stop
                                  str(stop_index),  # index of lien route element index
                                  str(1),  # enter
                                  str(1),  # exit
                                  time.strftime('%H:%M:%S', time.gmtime(arr)),  # arrival
                                  time.strftime('%H:%M:%S', time.gmtime(dep)) + '\n']))  # departure

    # vehicle journeys
    f.write(VisumTableHeader.table_vehicle_journeys_en)
    index = Visum.Net.AttValue(r"Max:VehJourneys\No")
    if index is not None:
        index += 1
    else:
        index = 1
    for line_key, line_val in lintim_files.lt_line.items():
        for lineroute_key, lineroute_val in line_val.items():
            for trips in lineroute_val['time_profile_trips_period']:
                f.write(';'.join([str(index),  # no
                                  time.strftime('%H:%M:%S', time.gmtime(trips)),
                                  str(line_key),  # line name
                                  str(lineroute_key),  # line route name
                                  lineroute_val['direction'],  # direction code
                                  '_'.join([str(line_key), str(lineroute_key)]),  # time profile name
                                  str(lineroute_val['time_profile_stop_order'][0]),
                                  # first time profile element index
                                  str(lineroute_val['time_profile_stop_order'][
                                          -1]) + '\n']))  # last time profile element index
                index += 1

    # vehicle journey items
    f.write(VisumTableHeader.table_vehicle_journey_items_en)
    index = Visum.Net.AttValue(r"Max:VehJourneys\No")
    if index is not None:
        index += 1
    else:
        index = 1
    for line_key, line_val in lintim_files.lt_line.items():
        for lineroute_key, lineroute_val in line_val.items():
            for trips in lineroute_val['time_profile_trips_period']:
                # double loop, list element index or each departure
                for idx in lineroute_val['time_profile_stop_order']:
                    f.write(';'.join([str(index),  # no
                                      str(idx) + '\n']))  # stop order index
                index += 1

    # vehicle journey sections
    f.write(VisumTableHeader.table_vehicle_journey_sections_en)
    index = Visum.Net.AttValue(r"Max:VehJourneys\No")
    if index is not None:
        index += 1
    else:
        index = 1
    for line_key, line_val in lintim_files.lt_line.items():
        for lineroute_key, lineroute_val in line_val.items():
            for trips in lineroute_val['time_profile_trips_period']:
                f.write(';'.join([str(index),  # vehicle journey items no
                                  str(1),   # vehicle journey sections no
                                  str(lineroute_val['time_profile_stop_order'][0]),
                                  # first time profile element index
                                  str(lineroute_val['time_profile_stop_order'][
                                          -1]) + '\n']))  # last time profile element index
                index += 1

    # if information exists, write lintim runtime as net attribute in visum
    if 'Overall runtime' in lintim_files.lt_values_dic_id:
        f.write(VisumTableHeader.table_network_lintim_en)  # Runtime LinTim
        f.write(lintim_files.lt_values_dic_id['Overall runtime'].replace('s','') + '\n')

# import net file to Visum
import_net2visum(path_net=sep_folder.join([folder_paths.network, name_put_supply_net]))

# set isroutepoint on all linerouteitems
Visum.Net.LineRouteItems.SetAllAttValues("ISROUTEPOINT", 1)






