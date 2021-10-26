import os

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


str_version = r"E:\Seafile\01_Interface\01_example\Version\example_stuttgart_teilnetz.ver"
Visum = open_visum(str_version)
sep_directory = "\\"

folder_path = Visum.UserPreferences.DocumentName.rsplit(sep_directory, 2)[
    0]  # path to opened Visum file (filenamepath), cut level to main folder level

name_instance = Visum.Net.AttValue("instance_name")
name_scenario = Visum.Net.AttValue("scenario_name")

folder_interface = sep_directory.join([folder_path, "_Interface"])
folder_visum2lintim = sep_directory.join([folder_interface, "Visum2Lintim"])
folder_lintim2visum = sep_directory.join([folder_interface, "Lintim2Visum"])
# folder_att = sep_directory.join([folder_visum2lintim, name_scenario + "_" + name_instance, "Attributes"])
# folder_net = sep_directory.join([folder_visum2lintim, name_scenario + "_" + name_instance, "Net"])

# create folder
# os.makedirs(folder_att, exist_ok=True)
# os.makedirs(folder_net, exist_ok=True)
os.makedirs(folder_lintim2visum, exist_ok=True)


a = 1