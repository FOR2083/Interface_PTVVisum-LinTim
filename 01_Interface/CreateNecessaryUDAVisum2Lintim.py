
# ====== Functions ========

def addStaticUDA( name="dummy", object="net", type="float"):
    # define visum type numbers of data types
    dict_types = {"str": 5,
                  "float": 2,
                  "int": 1,
                  "bool": 9}

    # type to type number
    typeno = dict_types[type.lower()]

    object = object[0].upper() + object[1:]

    # design of function call of ... .AddUserDefinedAttribute(id, short name, long name, value type)
    if object.lower() == "net":
        # case: net object
        str_method_call = "Visum.Net" + ".AddUserDefinedAttribute('" \
                          + name + "','" \
                          + name + "','" \
                          + name + "'," \
                          + str(typeno) + ")"
    else:
        # all other objects
        str_method_call = "Visum.Net." + object + ".AddUserDefinedAttribute('"\
                          + name + "','" \
                          + name + "','" \
                          + name + "'," \
                          + str(typeno) + ")"

    try:
        eval(str_method_call)
        print("UDA", name, object, " is created successfully")
    except:
        print("UDA", name, object, "is already defined")

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
# ====== Core ==========

str_version = r"E:\Seafile\01_Interface\01_example\Version\example_stuttgart_teilnetz.ver"
Visum = open_visum(str_version)

# insert UDA data here
# format:   [name, object, data type]
# e.g:      ["dummy", "links", "float"]
list_UDA = [["scenario_name", "Net", "str"],
            ["instance_name", "net", "str"],
            ["lintim_period_length", "net", "int"],
            ["lintim_time_units_per_minute", "net", "int"],
            ["lintim_base_unit_for_headway", "net", "int"],
            ["lintim_defdwelltime", "net", "int"],
            ["lintim_prepreptime", "net", "int"],
            ["lintim_postpreptime", "net", "int"],
            ["lintim_tsys_for_adapting", "Net", "str"],
            ["lintim_walktime_utility", "Net", "float"],
            ["lintim_transfer_utility", "Net", "int"],
            ["lintim_min_transfertime", "Net", "int"],
            ["internal_vol", "links", "float"],
            ["is_Terminal", "StopPoints", "bool"],
            ["is_Transfer", "StopPoints", "bool"]
            ]

for uda in list_UDA:
    addStaticUDA(uda[0], uda[1], uda[2])