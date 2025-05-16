import pandas as pd
import networkx as nx
from pathlib import Path


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
        Visum = com.Dispatch("Visum.Visum")
        print('open visum file: {}'.format(name))
        Visum.LoadVersion(name)
        print('erfolgreich geladen')

    return Visum


def calculate_runtime(Visum, tsys, attr_runtime_tsys_links, attr_runtime_tsys_connectors,
                      internal_runtime,
                      flag_runtime_between_nodes: bool=True,
                      flag_runtime_between_zones: bool=True):

    # Filter links with mode walk
    filter = Visum.Filters.LinkFilter()
    filter.Init()
    filter.AddCondition("OP_NONE", False, "TSysSet", "ContainsOneOf", tsys)
    filter.UseFilter = True

    attr_links = ["FromNodeNo", "ToNodeNo", attr_runtime_tsys_links]
    df_links = pd.DataFrame(Visum.Net.Links.GetMultipleAttributes(attr_links, OnlyActive=True), columns=attr_links,
                            dtype=int)

    # create graph
    G = nx.DiGraph()
    G.add_weighted_edges_from(df_links[["FromNodeNo", "ToNodeNo", attr_runtime_tsys_links]].values)

    # optional HERE: delete nodes without attribute, e. g. nodes that are not stop points
    # aggregate_edges(G, nodes_to_delete)

    # calculate shortest path for all nodes (only nodes of tsys links)
    iter_shortest_path_length = nx.algorithms.shortest_path_length(G, weight="weight")

    # write walking times in list: [FromNodeNo, ToNodeNo, ShortestPathWeight]
    walking_times = set()
    dict_shortest_path = dict()
    for from_node, inner_dict in iter_shortest_path_length:
        dict_shortest_path[from_node] = inner_dict
        for to_node, shortest_path_length in inner_dict.items():
            walking_times.add((from_node, to_node, shortest_path_length))
    df_runtime = pd.DataFrame(walking_times, columns=["FromNo", "ToNo", "runtime"])

    list_runtimes = []

    if flag_runtime_between_zones:
        # ass.: zone numbers don't interfere with node numbers
        attr_connectors = ["ZoneNo", "NodeNo", "Direction", attr_runtime_tsys_connectors]
        df_connectors = pd.DataFrame(Visum.Net.Connectors.GetMultipleAttributes(attr_connectors, OnlyActive=True),
                                     columns=attr_connectors, dtype=int)

        for direction, df in df_connectors.groupby("Direction"):
            if direction in ["Q", 1]:
                df_merge = df.merge(df_runtime, left_on="NodeNo", right_on="FromNo")
                df_q = pd.DataFrame(df_merge[["ZoneNo", "ToNo"]].values, columns=["FromNo", "ToNo"])
                df_q.loc[:, "runtime"] = df_merge["runtime"] + df_merge[attr_runtime_tsys_connectors]
                df_q = pd.concat([df_q, pd.DataFrame(df[["ZoneNo", "NodeNo", attr_runtime_tsys_connectors]].values,
                                                     columns=["FromNo", "ToNo", "runtime"])])
                df_q = df_q.groupby(["FromNo", "ToNo"]).agg(min).reset_index()
                list_runtimes.append(df_q)
            elif direction in ["Z", 2]:
                df_merge = df.merge(df_runtime, left_on="NodeNo", right_on="ToNo")
                df_z = pd.DataFrame(df_merge[["FromNo", "ZoneNo"]].values, columns=["FromNo", "ToNo"])
                df_z.loc[:, "runtime"] = df_merge["runtime"] + df_merge[attr_runtime_tsys_connectors]
                df_z = pd.concat([df_z, pd.DataFrame(df[["NodeNo", "ZoneNo", attr_runtime_tsys_connectors]].values,
                                                     columns=["FromNo", "ToNo", "runtime"])])
                df_z = df_z.groupby(["FromNo", "ToNo"]).agg(min).reset_index()
                list_runtimes.append(df_z)
            else:
                raise ValueError("Direction of connector is  not recognised")

        # Final: Merge  Z1 - N & N-Z2 to Z1 - Z2
        if 'df_q' in locals() and 'df_z' in locals():
            df_merge = df_q.merge(df_z, left_on="ToNo", right_on="FromNo", suffixes=("_q", "_z"))
            df_b = pd.DataFrame(df_merge[["FromNo_q", "ToNo_z"]].values, columns=["FromNo", "ToNo"])
            df_b.loc[:, "runtime"] = df_merge["runtime_q"] + df_merge["runtime_z"]
            list_runtimes.append(df_b)
        else:
            print("Fall nicht implementiert")

    if flag_runtime_between_nodes:
        list_runtimes.append(df_runtime)
        df_runtime = pd.concat(list_runtimes)
    else:
        df_runtime = pd.concat(list_runtimes)
        if flag_runtime_between_zones is False:
            raise ValueError("No runtime pairs active")
        else:
            set_zones = set(df_connectors["ZoneNo"].values.tolist())
            df_runtime.loc[df_runtime["FromNo"].isin(set_zones) & df_runtime["ToZone"].isin(set_zones), :]

    df_runtime = df_runtime.groupby(["FromNo", "ToNo"]).agg(min).reset_index()
    df_runtime.loc[df_runtime["FromNo"] == df_runtime["ToNo"], "runtime"] = internal_runtime

    return  df_runtime


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

# =========== CORE ===========================
str_version = r"E:\Seafile\01_Interface\01_example\Version\01_example.ver"


tsys = "Walk"
attr_runtime_tsys_links = "T_PuTSys(Walk)"
attr_runtime_tsys_connectors = "T0_TSys(Walk)"
internal_runtime = 0

# flag_runtime_between_zones:
# True = add walktimes/runtimes between all pair of zones
# False= only runtimes between nodes of graph
flag_runtime_between_zones = True
flag_runtime_between_nodes = True

Visum = open_visum(str_version)

name_scenario = Visum.Net.AttValue("scenario_name")

# path to Visum file (filenamepath), cut level to main folder level
folder_path = Path(Visum.UserPreferences.DocumentName).parents[1]

name_instance = Visum.Net.AttValue("instance_name")
name_scenario = Visum.Net.AttValue("scenario_name")

folder_interface = folder_path / "_Interface"
folder_visum2lintim = folder_interface / "Visum2Lintim"
folder_att = folder_visum2lintim / (name_scenario + "_" + name_instance)

output= folder_att / "walk_times.att"

df_walktime = calculate_runtime(Visum, tsys, attr_runtime_tsys_links, attr_runtime_tsys_connectors,
                      internal_runtime,
                      flag_runtime_between_nodes,
                      flag_runtime_between_zones)
df_walktime.rename(columns={
    "FromNo":"FromZoneNo",
    "ToNo": "ToZoneNo",
    "runtime":"Walktime_Minutes"
                }, inplace=True)

# change float to int
df_walktime["FromZoneNo"] = df_walktime["FromZoneNo"].astype(int)
df_walktime["ToZoneNo"] = df_walktime["ToZoneNo"].astype(int)

with open(output, "w", newline="\n") as f:
    write_object_to_net(object="OD pair",
                        df_object_attributes_to_write=df_walktime[["FromZoneNo", "ToZoneNo", "Walktime_Minutes"]],
                        file=f)
