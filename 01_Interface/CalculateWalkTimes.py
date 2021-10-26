import pandas as pd
import numpy as np
import networkx as nx


## code snippet: calculate walk times between stops + zones
# 1. create internal network (walk links only)
# 2. calculate shortest paths for all nodes in network
# 3. add connectors
# 4. return result as dataframe

# ========== FUNCTIONS =======================


def open_visum(name):
    try:
        ''' VisumStart
        - enables the usage in the procedure sequence or as Visum instance
        - name of the Visum version that is opened
        '''
        global Visum
        Visum
        name = Visum.UserPreferences.DocumentName
    except NameError:
        import win32com.client as com
        print('initialize visum instance')
        Visum = com.Dispatch("Visum.Visum.200")
        print('open visum file: {}'.format(name))
        Visum.LoadVersion(name)
        print('erfolgreich geladen')

    return Visum


## writes data of defined object type to .net file
# @param[in] f: target file opened in "write" or "append" mode
# @param[in] object: string of object type (as written in line "* Table: object" of net file )
# @param[in] df_object_attributes_to_write: DataFrame of object data to write. ! column names need to be names of visum attributes'
def write_object_to_net(object, df_object_attributes_to_write, file):
    '''
    Example
    attrNodes = ["No", "XCoord", "YCoord"]
    df_nodes = pandas.DataFrame(Visum.Net.Nodes.GetMultipleAttributes(attrNodes, OnlyActive=True), columns=attrNodes)

    with open("test.net", "w") as f:
        write_object_to_net("Node", df_nodes, f)
    '''

    header = ["*", "*"]
    header.insert(1, "* Table: " + object + "s")
    header.append(
        ("$" + object.upper().replace(" ", "") + ":" + ";".join(df_object_attributes_to_write.columns)).upper() + "\n")

    file.write("\n".join(header))
    df_object_attributes_to_write.to_csv(file, header=False, sep=";", index=False)


## consider zones for walk times
# @param[in] series_connector: connector data as type series
def add_walk_times_zones(series_connector):
    # todo: code creates duplicates
    walk_time_conn = series_connector["T0_TSys(Walk)"]
    zone_no = series_connector["ZoneNo"]
    from_node = series_connector["NodeNo"]

    if from_node in dict_shortest_path:
        # if node is part of link network
        # add walk time to zone to all nodes reached by from_node
        # walk time = existing walk time node pair + walk time connector

        # get nodes/walk time reached by from node
        inner_dict = dict_shortest_path[from_node]
        for to_node, shortest_path_length in inner_dict.items():
            # add walk times zone - to node & to node - zone

            if (zone_no, to_node) in walk_times:
                # if walk link exists -> choose minimum
                walk_times[(zone_no, to_node)] = np.min(
                    [shortest_path_length + walk_time_conn, walk_times[(zone_no, to_node)]])
                walk_times[(to_node, zone_no)] = np.min([shortest_path_length + walk_time_conn,
                                                         walk_times[(to_node, zone_no)]])
            else:
                walk_times[(zone_no, to_node)] = shortest_path_length + walk_time_conn
                walk_times[(to_node, zone_no)] = shortest_path_length + walk_time_conn

            if flag_walktime_between_zones == True:
                # add walk time to all zones of node
                # identify zones connected to to_node
                zones = df_Connectors.loc[
                    df_Connectors["NodeNo"] == to_node, ["ZoneNo", "T0_TSys(Walk)"]].values.tolist()
                for zone in zones:
                    to_zone = zone[0]
                    # walktime is sum of shortest path length and walktime of both connectors
                    walktime = shortest_path_length + walk_time_conn + zone[1]

                    # Case from_zone = to_zone --> walktime = 0
                    if zone_no == to_zone:
                        walktime = 0

                    # add walk time
                    if (zone_no, to_zone) not in walk_times:
                        walk_times[(zone_no, to_zone)] = walktime
                        walk_times[(to_zone, zone_no)] = walktime
                    else:
                        # if walk link exists -> choose minimum
                        walk_times[(zone_no, to_zone)] = np.min([walk_times[(zone_no, to_zone)], walktime])
                        walk_times[(to_zone, zone_no)] = np.min([walk_times[(to_zone, zone_no)], walktime])

    else:
        # if node not in link network
        # add
        if (zone_no, from_node) in walk_times:
            # if walk time exists -> choose minimum
            walk_times[(zone_no, from_node)] = np.min([walk_time_conn, walk_times[(zone_no, from_node)]])
            walk_times[(from_node, zone_no)] = np.min([walk_time_conn, walk_times[(zone_no, from_node)]])
        else:
            walk_times[(zone_no, from_node)] = walk_time_conn
            walk_times[(from_node, zone_no)] = walk_time_conn


# =========== CORE ===========================
str_version = r"E:\Seafile\01_Interface\01_example\Version\example_stuttgart_teilnetz.ver"
sep_directory = "\\"

# flag_walktime_between_zones:
# True = add walk times between all pair of zones
# False= for each zone add walk times to stops only
flag_walktime_between_zones = True

Visum = open_visum(str_version)

# path to opened Visum file (filenamepath), cut level to main folder level
folder_path = Visum.UserPreferences.DocumentName.rsplit(sep_directory, 2)[0]

name_instance = Visum.Net.AttValue("instance_name")
name_scenario = Visum.Net.AttValue("scenario_name")

folder_interface = sep_directory.join([folder_path, "_Interface"])
folder_visum2lintim = sep_directory.join([folder_interface, "Visum2Lintim"])
folder_att = sep_directory.join([folder_visum2lintim, name_scenario + "_" + name_instance])

output= sep_directory.join([folder_att,"walk_times.att"])

#str_output_walk_times = Visum.GetPath(15) + "walk_times.att"

# Filter links with mode walk
filter = Visum.Filters.LinkFilter()
filter.Init()
filter.AddCondition("OP_NONE", False, "TSysSet", "ContainsOneOf", "Walk")
filter.UseFilter = True

attrLinks = ["FromNodeNo", "ToNodeNo", "T_PuTSys(Walk)"]
df_Links = pd.DataFrame(Visum.Net.Links.GetMultipleAttributes(attrLinks, OnlyActive=True), columns=attrLinks)

# seconds to minute
df_Links["T_PuTSys(Walk)"] = df_Links["T_PuTSys(Walk)"] / 60

# create graph
G = nx.Graph()
G.add_weighted_edges_from(df_Links[["FromNodeNo", "ToNodeNo", "T_PuTSys(Walk)"]].values)

# if zone != stop -> add connectors as additional edges
# ass.: zone numbers don't interfere with node numbers
attrConnectors = ["ZoneNo", "NodeNo", "T0_TSys(Walk)"]
df_Connectors = pd.DataFrame(Visum.Net.Connectors.GetMultipleAttributes(attrConnectors, OnlyActive=True),
                             columns=attrConnectors)

# seconds to minute
df_Connectors["T0_TSys(Walk)"] = df_Connectors["T0_TSys(Walk)"] / 60

# calculate shortest path for all nodes (only nodes of walk links)
iter_shortest_path_length = nx.algorithms.shortest_path_length(G, weight="weight")

# write walking times in list: [FromNodeNo, ToNodeNo, ShortestPathWeight]
walk_times = dict()
dict_shortest_path = dict()
for from_node, inner_dict in iter_shortest_path_length:
    dict_shortest_path[from_node] = inner_dict
    for to_node, shortest_path_length in inner_dict.items():
        walk_times[(from_node, to_node)] = shortest_path_length

df_Connectors.apply(add_walk_times_zones, axis=1)

# add walk time from node to stop with same number (0s)
for node in G.nodes:
    walk_times[(node, node)] = 0

df_walktime = pd.DataFrame(
    data={"Walktime_Minutes": list(walk_times.values())},
    index=pd.MultiIndex.from_tuples(tuples=walk_times.keys(), names=["FromZoneNo", "ToZoneNo"])
)

df_walktime["FromZoneNo"] = df_walktime.index.get_level_values(0)
df_walktime["ToZoneNo"] = df_walktime.index.get_level_values(1)

# test
walk_time_matrix = df_walktime["Walktime_Minutes"].unstack()
if walk_time_matrix.shape[0] != walk_time_matrix.shape[1]:
    ValueError("walk time matrix is not quadratic")

# check if every zone has at least one connected stop/valid walk time
list_zones = list(Visum.Net.Zones.GetMultipleAttributes(["No"], OnlyActive=False))
list_zones = [float(x[0]) for x in list_zones]
min_walk_time_zone = walk_time_matrix.loc[
    walk_time_matrix.index.isin(list_zones), ~walk_time_matrix.columns.isin(list_zones)].min(axis=1)

invalid = min_walk_time_zone.loc[min_walk_time_zone.isna()].index
if len(invalid) > 0:
    ValueError("the following zones are not connected: " + ",".join(invalid.astype(str)))

# df_walktime.sort_values(by=["Walktime_Minutes"], inplace=True)
# df_walktime.sort_values(by=["FromZoneNo", "ToZoneNo"], inplace=True)

# change float to int
df_walktime["FromZoneNo"] = df_walktime["FromZoneNo"].astype(int)
df_walktime["ToZoneNo"] = df_walktime["ToZoneNo"].astype(int)

with open(output, "w", newline="\n") as f:
    write_object_to_net(object="OD pair",
                        df_object_attributes_to_write=df_walktime[["FromZoneNo", "ToZoneNo", "Walktime_Minutes"]],
                        file=f)
