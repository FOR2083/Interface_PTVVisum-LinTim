import pandas as pd
import win32com.client as com
import itertools as it
import networkx as nx
import numpy as np

# =========== Functions ======================
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
        print('initialize visum instance')
        Visum = com.Dispatch("Visum.Visum.200")
        print('open visum file: {}'.format(name))
        Visum.LoadVersion(name)
        print('erfolgreich geladen')

    return Visum

def aggregate_links(links_tsys, nodes, tsys ):
    str_t = "T_PuTSys(" + tsys + ")"

    subgraph_tsys = nx.DiGraph()
    subgraph_tsys.add_weighted_edges_from(links_tsys.loc[:, ["FromNodeNo", "ToNodeNo", str_t]].values)
    edge_attr= { (x,y): value for x,y,value in links_tsys.loc[:, ["FromNodeNo", "ToNodeNo", "Length"]].values}
    nx.set_edge_attributes(subgraph_tsys, edge_attr, "length")

    # init
    node_attr = {no: "" for no, node_data in subgraph_tsys.nodes(data=True) }
    nx.set_node_attributes(subgraph_tsys, node_attr, "TSysSet")

    node_attr = {nodes.loc[index, "No"]: nodes.loc[index, "TSysSet"] for index in nodes.index}
    nx.set_node_attributes(subgraph_tsys, node_attr, "TSysSet")

    nodes_to_delete = [no for no, node_data in subgraph_tsys.nodes(data=True) if tsys not in node_data["TSysSet"]]
    for node in nodes_to_delete:
        keys_new_edges = [(x, y) for x, y in it.product(
            subgraph_tsys.predecessors(node),
            subgraph_tsys.successors(node)) if x != y]
        weights = [[x, y, sum([subgraph_tsys[x][node]["weight"],
                               subgraph_tsys[node][y]["weight"]])] for x, y in keys_new_edges]
        length = {(x, y): sum([subgraph_tsys[x][node]["length"],
                               subgraph_tsys[node][y]["length"]]) for x, y in keys_new_edges}
        subgraph_tsys.add_weighted_edges_from(weights)
        nx.set_edge_attributes(subgraph_tsys, length, "length")
        subgraph_tsys.remove_node(node)

    return subgraph_tsys

# merge link infos
def merge_link_data(link_data, tsys):
    key = (link_data["FromNodeNo"], link_data["ToNodeNo"])
    str_t = "T_PuTSys(" + tsys + ")"

    if key in dict_links:
        dict_links[key]["TSysSet"].add(tsys)
    else:
        # init link
        dict_links[key] = {"FromNodeNo": key[0],
                          "Length":link_data["length"],
                            "ToNodeNo": key[1],
                            "TSysSet": set([tsys]),
                           "T_PuTSys(B)": 9999999,
                           "T_PuTSys(U)": 9999999,
                           "T_PuTSys(LRT)": 9999999,
                           "T_PuTSys(HRT)": 9999999,
                           "T_PuTSys(Walk)": 9999999
                           }
    # set t of tsys
    dict_links[key][str_t] = link_data["weight"]

## Schreibt den Header einer .net Datei
# @param [in] f: geöffnete Datei, in die die Informationen geschrieben werden
def write_header_net(f):
    header = '''$VISION
* Universitaet Stuttgart Fakultaet 2 Bau+Umweltingenieurwissenschaften Stuttgart
* 28.03.16
*
* Tabelle: Versionsblock
*
$VERSION:VERSNR;FILETYPE;LANGUAGE;UNIT
10;Net;DEU;KM
'''
    #f = open(path_net + name_net, "a")
    f.write(header)

## Schreibt Knoten in eine .net Datei
# @param [in] f: geöffnete Datei, in die die Informationen geschrieben werden
# @param [in] list_nodes: Liste mit Knotendaten
def write_nodes_to_net(list_nodes, f):
    # header .net file
    header_nodes = '''
* 
* Table: Nodes
* 
$NODE:NO;XCOORD;YCOORD
'''

    # convert inner list to string
    list_str = [';'.join([str(int(lst[0])), str(lst[1]), str(lst[2])]) for lst in list_nodes]

    #f = open(path_net + name_net, mode)
    # f.write(header)

    # write nodes
    f.write(header_nodes)
    f.write("\n".join(list_str ))

## Schreibt Bezirke in eine .net Datei
# @param [in] f: geöffnete Datei, in die die Informationen geschrieben werden
# @param [in] list_zones: Liste mit Bezirksdaten
def write_zones_to_net(list_zones, f):
    # header .net file
    header_nodes = '''
* 
* Table: Zones
* 
$ZONE:NO;XCOORD;YCOORD
'''

    # convert inner list to string
    list_str = [';'.join([str(int(lst[0])), str(lst[1]), str(lst[2])]) for lst in list_zones]

    #f = open(path_net + name_net, mode)
    # f.write(header)

    # write nodes
    f.write(header_nodes)
    f.write("\n".join(list_str ))

## Schreibt Haltepunkte in eine .net Datei
# @param [in] f: geöffnete Datei, in die die Informationen geschrieben werden
# @param [in] list_stoppoints: Liste mit Haltepunktdaten
def write_stoppoints_to_net(list_stoppoints, f):
    header_stoppoints = '''
* 
* Table: Stoppoints
* 
$STOPPOINT:NO;NODENO;TSYSSET;DEFDWELLTIME;IS_TERMINAL;IS_TRANSFER
'''
    # convert inner list to string
    list_str = [';'.join([str(int(lst[0])), str(int(lst[1])), str(lst[2]), lst[3], str(lst[4]), str(lst[5])]) for lst in list_stoppoints]

    # f = open(path_net + name_net, mode)

    # write stops
    f.write(header_stoppoints)
    f.write("\n".join(list_str))

## Schreibt Strecken in eine .net Datei
# wandelt Streckeninformationen in das benötigte Format um und schreibt sie in die Datei
# @param [in] f: geöffnete Datei, in die die Informationen geschrieben werden
# @param [in] links: dataframe mit Streckendaten
def write_links_to_net_runtimes(links, f):
    text = []
    for key, link in links.iterrows():
        tsysset = ",".join(link["TSysSet"])

        # write link for both directions in net file
        # first direction
        text.append((str(int(link["FromNodeNo"])) + ";" +
                    str(int(link["ToNodeNo"])) + ";" +
                    tsysset + ";" +
                    str(link["Length"]) + "km;" +
                    str(int(link["T_PuTSys(B)"])) + "s;" +
                    str(int(link["T_PuTSys(U)"])) + "s;" +
                    str(int(link["T_PuTSys(LRT)"])) + "s;" +
                    str(int(link["T_PuTSys(HRT)"])) + "s;" +
                     str(int(link["T_PuTSys(Walk)"])) + "s").replace("9999999s","") + ";" +
                    str(link["Internal_Vol"])) # todo str(link["internal_vol"]))    # internal volume

    header_links = '''
* 
* Table: Links
* 
$LINK:FROMNODENO;TONODENO;TSYSSET;LENGTH;t_PuTSys(B);t_PuTSys(U);t_PuTSys(LRT);t_PuTSys(HRT);T_PuTSys(Walk);INTERNAL_VOL
'''
    #
    #f = open(path_net + name_net, mode)
    #f.write(header)
    f.write(header_links)
    f.write("\n".join(text))

def write_links_to_net(links, f):
    header_links = '''
* 
* Table: Links
* 
$LINK:NO;FROMNODENO;TONODENO;
'''
    text = []
    for key, link in links.iterrows():
        # write link for both directions in net file
        # first direction
        text.append(str(int(link["No"])) + ";" +
                    str(int(link["FromNodeNo"])) + ";" +
                    str(int(link["ToNodeNo"]))
                    )

    f.write(header_links)
    f.write("\n".join(text))

def write_lines(f):
    header_lines ='''
* 
* Table: Lines
* 
$LINE:NAME;TSYSCODE
'''
    lines = list(Visum.Net.Lines.GetMultipleAttributes(["Name", "TSysCode"], OnlyActive=True))
    lines = [";".join(line) for line in lines]
    f.write(header_lines)
    f.write("\n".join(lines))

def write_line_routes(f):
    header_line_routes = '''
* 
* Table: Line routes
* 
$LINEROUTE:LINENAME;NAME;DIRECTIONCODE;ISCIRCLELINE    
'''
    lineroutes = list(Visum.Net.LineRoutes.GetMultipleAttributes(["LINENAME","NAME","DIRECTIONCODE","ISCIRCLELINE"], OnlyActive=True))
    lineroutes = [";".join([line[0], line[1], line[2], str(int(line[3]))]) for line in lineroutes]
    f.write(header_line_routes)
    f.write("\n".join(lineroutes))

def write_line_route_items(f):
    header_line_route_items = '''
* 
* Table: Line route items
* 
$LINEROUTEITEM:LINENAME;LINEROUTENAME;DIRECTIONCODE;INDEX;ISROUTEPOINT;NODENO;STOPPOINTNO
'''
    linerouteitems = list(Visum.Net.LineRouteItems.GetMultipleAttributes(["LINENAME","LINEROUTENAME","DIRECTIONCODE","INDEX","ISROUTEPOINT","NODENO","STOPPOINTNO"], OnlyActive=True))
    linerouteitems = [";".join([str(item) for item in lritem]) for lritem in linerouteitems]

    # remove decimal place of numbers
    linerouteitems = [item.replace(".0", "") for item in linerouteitems]

    # write
    f.write(header_line_route_items)
    f.write("\n".join(linerouteitems))

def write_veh_journeys(f):
    header_veh_journeys = '''
* 
* Table: Vehicle journeys
* 
$VEHJOURNEY:NO;DEP;LINENAME;LINEROUTENAME;DIRECTIONCODE;TIMEPROFILENAME;FROMTPROFITEMINDEX;TOTPROFITEMINDEX
'''
    veh_journeys = pd.DataFrame(Visum.Net.VehicleJourneys.GetMultipleAttributes(["NO","DEP","LINENAME","LINEROUTENAME","DIRECTIONCODE","TIMEPROFILENAME","FROMTPROFITEMINDEX","TOTPROFITEMINDEX"], OnlyActive=True))
    # convert time
    veh_journeys[1] = pd.to_datetime(veh_journeys[1], unit="s").dt.time
    # convert to list of strings
    veh_journeys = veh_journeys[:].astype(str).values.tolist()
    # remove decimal place of numbers
    veh_journeys = [";".join(item) for item in veh_journeys]
    veh_journeys = [item.replace(".0", "") for item in veh_journeys]
    f.write(header_veh_journeys)
    f.write("\n".join(veh_journeys))

def write_time_profiles(f):
    header_time_profiles = '''
* 
* Table: Time profiles
* 
$TIMEPROFILE:LINENAME;LINEROUTENAME;NAME;DIRECTIONCODE   
'''
    tps = list(Visum.Net.TimeProfiles.GetMultipleAttributes(["LINENAME", "LINEROUTENAME", "NAME", "DIRECTIONCODE"],
                                                                 OnlyActive=True))
    tps = [";".join(tp) for tp in tps]
    f.write(header_time_profiles)
    f.write("\n".join(tps))

def write_time_profile_items(f):
    header_time_profile_items = '''
* 
* Table: Time profile items
* 
$TIMEPROFILEITEM:LINENAME;LINEROUTENAME;TIMEPROFILENAME;DIRECTIONCODE;ARR;DEP;
'''
    timeprofiles = pd.DataFrame(Visum.Net.TimeProfileItems.GetMultipleAttributes(["LINENAME","LINEROUTENAME","TIMEPROFILENAME", "DIRECTIONCODE","ARR", "DEP",], OnlyActive=True))
    # convert time
    timeprofiles[4] = pd.to_datetime(timeprofiles[4], unit="s").dt.time
    timeprofiles[5] = pd.to_datetime(timeprofiles[5], unit="s").dt.time
    # convert to list of strings
    timeprofiles = timeprofiles[:].astype(str).values.tolist()
    # remove decimal place of numbers
    timeprofiles = [";".join(item) for item in timeprofiles]
    timeprofiles = [item.replace(".0", "") for item in timeprofiles]
    f.write(header_time_profile_items)
    f.write("\n".join(timeprofiles))

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
    header.append(("$" + object.upper().replace(" ","") + ":" + ";".join(df_object_attributes_to_write.columns)).upper() + "\n")

    file.write("\n".join(header))
    df_object_attributes_to_write.to_csv(file, header=False, sep=";", index=False)

# =========== CORE ===========================
str_version = r"E:\Seafile\36_VRS\Version\36_VRS_LHS_SR2.ver"
sep_directory = "\\"

# flag_walktime_between_zones:
# True = add walktimes between all pair of zones
# False= for each zone add walktimes to stops only
flag_walktime_between_zones = True

Visum = open_visum(str_version)

# path to opened Visum file (filenamepath), cut level to main folder level
folder_path = Visum.UserPreferences.DocumentName.rsplit(sep_directory, 2)[
    0]

name_instance = Visum.Net.AttValue("instance_name")
name_scenario = Visum.Net.AttValue("scenario_name")
folder_instance = sep_directory.join([folder_path, "_Interface", "Visum2Lintim", name_scenario + "_" + name_instance])

str_output_infrastructure = sep_directory.join([folder_instance, "infrastructure.net"])
str_output_fixed_lines =sep_directory.join([folder_instance, "visum-fixed-lines.net"])
str_output_fixed_times =sep_directory.join([folder_instance, "visum-fixed-times.net"])
#str_output_walk_times = sep_directory.join([folder_instance, "Attributes", "walk_times.att"])

#pd.options.display.float_format = '${:, .2f}'.format

# init
attr_links = ["FromNodeNo", "Length", "ToNodeNo", "TSysSet", "T_PuTSys(B)", "T_PuTSys(U)", "T_PuTSys(LRT)", "T_PuTSys(HRT)", "T_PuTSys(Walk)" , "internal_vol"]
col_dtype = {"FromNodeNo": int, "Length": float, "ToNodeNo": int, "TSysSet": object, "T_PuTSys(B)": int, "T_PuTSys(U)": int, "T_PuTSys(LRT)": int, "T_PuTSys(HRT)": int, "T_PuTSys(Walk)": int}
df_links = pd.DataFrame(Visum.Net.Links.GetMultipleAttributes(attr_links, OnlyActive=False), columns=attr_links)#, dtype=col_dtype)
df_links = df_links.astype(col_dtype)
#df_links["TSysSet"] = df_links["TSysSet"].apply(lambda x: set(x.split(",")))

attr_stop_points = ["No", "NodeNo", "TSysSet", "DefDwellTime", "Is_Terminal", "Is_Transfer"]
df_stop_points = pd.DataFrame(Visum.Net.StopPoints.GetMultipleAttributes(attr_stop_points, OnlyActive=False), columns=attr_stop_points)
df_stop_points["TSysString"] = df_stop_points["TSysSet"]
df_stop_points["TSysSet"] = df_stop_points["TSysSet"].apply(lambda x: set(x.split(",")))
df_stop_points["DefDwellTime"] = df_stop_points["DefDwellTime"].apply(lambda x: str(int(x)) + "s")
df_stop_points.astype({"No": int, "NodeNo": int})

#list_tsys = Visum.Net.AttValue("tsys")
#list_tsys.split(",")
list_tsys = ["B", "LRT", "U", "HRT"]
dict_g_tsys = {}

# eliminate nodes without stop for each transport system
for tsys in list_tsys:
    if tsys == "Walk":
        continue

    # dataframe with links for tsys
    df_tsys = df_links[df_links["TSysSet"].str.contains(tsys)]

    # skip tsys without links
    if len(df_tsys) < 1:
        continue
    # create network of tsys, delete nodes which arent't stops for given tsys
    dg_tsys = aggregate_links(links_tsys=df_tsys, nodes=df_stop_points, tsys=tsys)

    # store tsys, network in dict
    dict_g_tsys[tsys] = dg_tsys

dict_links = {}
# get data of networks
for tsys, dg_tsys in dict_g_tsys.items():
    edges = nx.to_pandas_edgelist(dg_tsys, source="FromNodeNo", target="ToNodeNo")
    edges.apply(lambda x: merge_link_data(link_data=x, tsys=tsys),axis=1)

# add walk links
tsys = "Walk"
edges = df_links.loc[df_links["TSysSet"].str.contains(tsys), ["FromNodeNo", "ToNodeNo", "T_PuTSys(Walk)", "Length"]]
edges["weight"] = edges["T_PuTSys(Walk)"]
edges["length"] = edges["Length"]
edges.apply(lambda x: merge_link_data(link_data=x, tsys=tsys),axis=1)

# dictionary to dataframe
df_new_links = pd.DataFrame.from_dict(dict_links, orient="index")

# add link no
df_new_links["OuterNodes"] = tuple(zip(np.min(df_new_links[["FromNodeNo", "ToNodeNo"]], axis= 1),np.max(df_new_links[["FromNodeNo", "ToNodeNo"]], axis=1)))
unique_links = np.unique(df_new_links["OuterNodes"].values)
link_no = dict(zip(unique_links, range(len(unique_links))))

def create_link_no(outer_nodes):
    return link_no[outer_nodes] + 1

df_new_links["No"] = df_new_links["OuterNodes"].apply(create_link_no)

# merge internal volume (got lost during aggregation)
def get_internal_volume(series):
    idx = series.name
    if idx in df_links.index:
        return df_links.loc[idx, "internal_vol"]
    else:
        return 0

df_links.index = tuple(zip(df_links["FromNodeNo"], df_links["ToNodeNo"]))
df_new_links.loc[:, "Internal_Vol"] = df_new_links.apply(get_internal_volume, axis=1)
df_new_links.loc[:,"Internal_Vol"].replace(np.nan, 0, inplace=True)

# import visum infrastructure:
attr_nodes = ["No", "XCOORD", "YCOORD"]
df_nodes = pd.DataFrame(Visum.Net.Nodes.GetMultipleAttributes(attr_nodes, OnlyActive=False), columns=attr_nodes)

attr_zones = ["No", "XCOORD", "YCOORD"]
df_zones = pd.DataFrame(Visum.Net.Zones.GetMultipleAttributes(attr_zones, OnlyActive=False), columns=attr_zones)

# check if nodes & zones have unique numbers
set_zones = set(df_zones["No"])
set_nodes = set(df_nodes["No"])

# if len(set_zones.intersection(set_nodes)) > 0:
#     print(' Change numbers of nodes and zones')
#     raise Exception('zone and node identifier are not unique')

# write infrastructure.net file
with open(str_output_infrastructure, "w") as f:
    write_header_net(f)
    write_nodes_to_net(df_nodes.values.tolist(),f)
    write_zones_to_net(df_zones.values.tolist(),f)
    write_links_to_net_runtimes(df_new_links, f)
    write_stoppoints_to_net(df_stop_points.loc[:,["No", "NodeNo", "TSysString", "DefDwellTime", "Is_Terminal", "Is_Transfer"]].values.tolist(),f)

if Visum.Net.Lines.Count > 0:
    # write visum-fixed-lines.net
    with open(str_output_fixed_lines, "w") as f:
        write_header_net(f)
        write_links_to_net(df_new_links[["No", "FromNodeNo", "ToNodeNo"]], f)
        write_lines(f)
        write_line_routes(f)
        write_line_route_items(f)
        # for lintim to get frequencies
        write_veh_journeys(f)

    if Visum.Net.TimeProfiles.Count > 0:
        # write visum-fixed-times.net
        with open(str_output_fixed_times, "w") as f:
            write_header_net(f)
            write_links_to_net(df_new_links[["No", "FromNodeNo", "ToNodeNo"]], f)
            write_lines(f)
            write_line_routes(f)
            write_line_route_items(f)
            write_time_profiles(f)
            write_time_profile_items(f)
            write_veh_journeys(f)

del Visum
