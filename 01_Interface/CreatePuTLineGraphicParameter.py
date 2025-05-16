'''
OUTPUT:         GraphicParameterPerLine.gpax
DESCRIPTION:    creates a Visum graphic parameter file to visualize
                the line routes. For each existing line in visum  graphic parameters are created
'''

class VisumFolderStructure:
    def __init__(self, Visum, sep_folder):
        self.main = Visum.UserPreferences.DocumentName.rsplit(sep_folder, 2)[0]
        self.version = Visum.GetPath(2)
        self.network = Visum.GetPath(1)
        self.graphicparameter = Visum.GetPath(8)


class VisumLineRoute:
    def __init__(self):
        self.header = ''.join([0 * '\t' + '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + '\n',
                               0 * '\t' + '<networkGraphicParameters version="1">' + '\n',
                               1 * '\t' + '<layerParameters>' + '\n',
                               2 * '\t' + '<layers partial="true">' + '\n',
                               3 * '\t' + '<layer draw="true" type="ROUTECOURSES"/>' + '\n',
                               2 * '\t' + '</layers>' + '\n',
                               1 * '\t' + '</layerParameters>' + '\n',
                               1 * '\t' + '<routeCourses aggrType="LINE" drawOnlyActiveLineRoutes="false" intersectionPreference="PREFERINTERSECTINGLINESATNODESWITHMULTIPLELEGS">' + '\n',
                               2 * '\t' + '<display>' + '\n',
                               3 * '\t' + '<lines>' + '\n'])
        self.footer = ''.join([3 * '\t' + '</lines>' + '\n',
                               2 * '\t' + '</display>' + '\n',
                               2 * '\t' + '<label>' + '\n',
                               3 * '\t' + '<transportSystems adoptColorFromLineStyle="true" draw="false">' + '\n',
                               4 * '\t' + '<text attrID="NAME" drawFrame="true" frameColor="ff000000" numDecPlaces="4" roundValue="1" textColor="ff000000" textSize="1.8" transparent="false">' + '\n',
                               5 * '\t' + '<fontStyle bold="false" fontFamilyName="Arial" italic="false"/>' + '\n',
                               4 * '\t' + '</text>' + '\n',
                               3 * '\t' + '</transportSystems>' + '\n',
                               3 * '\t' + '<mainLines adoptColorFromLineStyle="true" draw="false">' + '\n',
                               4 * '\t' + '<text attrID="NAME" drawFrame="true" frameColor="ff000000" numDecPlaces="4" roundValue="1" textColor="ff000000" textSize="1.8" transparent="false">' + '\n',
                               5 * '\t' + '<fontStyle bold="false" fontFamilyName="Arial" italic="false"/>' + '\n',
                               4 * '\t' + '</text>' + '\n',
                               3 * '\t' + '</mainLines>' + '\n',
                               3 * '\t' + '<lines adoptColorFromLineStyle="false" draw="true">' + '\n',
                               4 * '\t' + '<text attrID="LABEL" drawFrame="true" frameColor="ff000000" numDecPlaces="0" roundValue="1" textColor="ff000000" textSize="3" transparent="false">' + '\n',
                               5 * '\t' + '<fontStyle bold="false" fontFamilyName="Arial" italic="false"/>' + '\n',
                               3 * '\t' + '</text>' + '\n',
                               3 * '\t' + '</lines>' + '\n',
                               3 * '\t' + '<lineRoutes adoptColorFromLineStyle="true" draw="false">' + '\n',
                               4 * '\t' + '<text attrID="NAME" drawFrame="true" frameColor="ff000000" numDecPlaces="4" roundValue="1" textColor="ff000000" textSize="1.8" transparent="false">' + '\n',
                               5 * '\t' + '<fontStyle bold="false" fontFamilyName="Arial" italic="false"/>' + '\n',
                               4 * '\t' + '</text>' + '\n',
                               3 * '\t' + '</lineRoutes>' + '\n',
                               3 * '\t' + '<operators adoptColorFromLineStyle="true" draw="false">' + '\n',
                               4 * '\t' + '<text attrID="NAME" drawFrame="true" frameColor="ff000000" numDecPlaces="4" roundValue="1" textColor="ff000000" textSize="1.8" transparent="false">' + '\n',
                               5 * '\t' + '<fontStyle bold="false" fontFamilyName="Arial" italic="false"/>' + '\n',
                               4 * '\t' + '</text>' + '\n',
                               3 * '\t' + '</operators>' + '\n',
                               2 * '\t' + '</label>' + '\n',
                               1 * '\t' + '</routeCourses>' + '\n',
                               0 * '\t' + '</networkGraphicParameters>'])
        self.net_obj = 4 * '\t' + '<netObj draw="true" name="'
        self.net_obj_closer = 4 * '\t' + '</netObj>' + '\n'
        self.line_style = 5 * '\t' + '<lineStyle>' + '\n'
        self.line_style_closer = 5 * '\t' + '</lineStyle>' + '\n'
        self.stokes = 6 * '\t' + '<strokes>' + '\n'
        self.stokes_closer = 6 * '\t' + '</strokes>' + '\n'
        self.stoke_color = 7 * '\t' + '<stroke color="'
        self.dash_pattern = 'dashPattern="'
        self.dash_width = 'dashWidth="'
        self.width = 'width="'
        self.closer = '>'
        self.closer_slash = '/>'

        '''
        dyle lot (Farbrad) CYMK-values https://s-media-cache-ak0.pinimg.com/originals/98/a3/ea/98a3ea9aeffcfb009d3ce46da592dabd.jpg
        index = color column
        light = last entry in color description e.g. 40.100.0.60, 60 = light

        alte color codes
        self.line_route_colors_in = [['ff0000',     # index: 1,     light: 0,   CYMK:   0   100     100     0
                                      'FF5900',     # index: 3,     light: 0,   CYMK:   0   65      100     0
                                      'FFFF00',     # index: 5,     light: 0,   CYMK:   0   0       100     0
                                      '00FF19',     # index: 7,     light: 0,   CYMK:   100 0       90      0
                                      '0019FF',     # index: 9,     light: 0,   CYMK:   100 60      0       0
                                      '9900FF',     # index: 11,    light: 0,   CYMK:   80  100     0       0

                                      '8C0000',     # index: 1,     light: 30,  CYMK:   0   100     100     30
                                      '8C0E1C',     # index: 3,     light: 30,  CYMK:   0   65      100     30
                                      '8C3100',     # index: 5,     light: 30,  CYMK:   0   0       100     30
                                      '8C5400',     # index: 7,     light: 30,  CYMK:   100 0       90      30
                                      '8C8C00',     # index: 9,     light: 30,  CYMK:   100 60      0       30
                                      '388C00',     # index: 11,    light: 30,  CYMK:   80  100     0       30

                                      '008C0E',     # index: 1,     light: 60,  CYMK:   0   100     100     60
                                      '008C54',     # index: 3,     light: 60,  CYMK:   0   65      100     60
                                      '00388C',     # index: 5,     light: 60,  CYMK:   0   0       100     60
                                      '000E8C',     # index: 7,     light: 60,  CYMK:   100 0       90      60
                                      '1C008C',     # index: 9,     light: 60,  CYMK:   100 60      0       60
                                      '54008C',     # index: 11,    light: 60,  CYMK:   80  100     0       60

                                      'D90000',     # index: 2,     light: 0,   CYMK:   0   90      80      0
                                      'D9162B',     # index: 4,     light: 0,   CYMK:   0   40      100     0
                                      'D94C00',     # index: 6,     light: 0,   CYMK:   60  0       100     0
                                      'D98200',     # index: 8,     light: 0,   CYMK:   100 0       40      0
                                      'D9D900',     # index: 10,    light: 0,   CYMK:   100 90      0       0
                                      '57D900',     # index: 12,    light: 0,   CYMK:   40  100     0       0

                                      'D90000',     # index: 2,     light: 30,  CYMK:   0   90      80      30
                                      'D9162B',     # index: 4,     light: 30,  CYMK:   0   40      100     30
                                      'D94C00',     # index: 6,     light: 30,  CYMK:   60  0       100     30
                                      'D98200',     # index: 8,     light: 30,  CYMK:   100 0       40      30
                                      'D9D900',     # index: 10,    light: 30,  CYMK:   100 90      0       30
                                      '57D900',     # index: 12,    light: 30,  CYMK:   40  100     0       30

                                      '00D916',     # index: 2,     light: 60,  CYMK:   0   90      80      60
                                      '00D982',     # index: 4,     light: 60,  CYMK:   0   40      100     60
                                      '0057D9',     # index: 6,     light: 60,  CYMK:   60  0       100     60
                                      '0016D9',     # index: 8,     light: 60,  CYMK:   100 0       40      60
                                      '2B00D9',     # index: 10,    light: 60,  CYMK:   100 90      0       60
                                      '8200D9',     # index: 12,    light: 60,  CYMK:   40  100     0       60

                                      'FF0000',     # index: 1,     light: 15,  CYMK:   0   100     100     15
                                      'FF5900',     # index: 3,     light: 15,  CYMK:   0   65      100     15
                                      'FFFF00',     # index: 5,     light: 15,  CYMK:   0   0       100     15
                                      '00FF19',     # index: 7,     light: 15,  CYMK:   100 0       90      15
                                      '0019FF',     # index: 9,     light: 15,  CYMK:   100 60      0       15
                                      '9900FF',     # index: 11,    light: 15,  CYMK:   80  100     0       15

                                      'FF0000',     # index: 1,     light: 45,  CYMK:   0   100     100     45
                                      'FF5900',     # index: 3,     light: 45,  CYMK:   0   65      100     45
                                      'FFFF00',     # index: 5,     light: 45,  CYMK:   0   0       100     45
                                      '00FF19',     # index: 7,     light: 45,  CYMK:   100 0       90      45
                                      '0019FF',     # index: 9,     light: 45,  CYMK:   100 60      0       45
                                      '9900FF',     # index: 11,    light: 45,  CYMK:   80  100     0       45

                                      'FF0000',     # index: 2,     light: 15,  CYMK:   0   90      80      15
                                      'FF5900',     # index: 4,     light: 15,  CYMK:   0   40      100     15
                                      'FFFF00',     # index: 6,     light: 15,  CYMK:   60  0       100     15
                                      '00FF19',     # index: 8,     light: 15,  CYMK:   100 0       40      15
                                      '0019FF',     # index: 10,    light: 15,  CYMK:   100 90      0       15
                                      '9900FF',     # index: 12,    light: 15,  CYMK:   40  100     0       15

                                      'FF0000',     # index: 2,     light: 45,  CYMK:   0   90      80      45
                                      'FF5900',     # index: 4,     light: 45,  CYMK:   0   40      100     45
                                      'FFFF00',     # index: 6,     light: 45,  CYMK:   60  0       100     45
                                      '00FF19',     # index: 8,     light: 45,  CYMK:   100 0       40      45
                                      '0019FF',     # index: 10,    light: 45,  CYMK:   100 90      0       45
                                      '0019FF',     # index: 12,    light: 45,  CYMK:   40  100     0       45

                                      'FF0000',     # index: 1,     light: -1,  CYMK:   0   85      75      0
                                      'FF5900',     # index: 3,     light: -1,  CYMK:   0   50      80      0
                                      'FFFF00',     # index: 5,     light: -1,  CYMK:   0   0       0       80
                                      '00FF19',     # index: 7,     light: -1,  CYMK:   0   100     100     0
                                      '0019FF',     # index: 9,     light: -1,  CYMK:   85  50      0       0
                                      '9900FF',     # index: 11,    light: -1,  CYMK:   65  85      0       0

                                      'FF0000',     # index: 1,     light: -3,  CYMK:   0   45      30      0
                                      'FF5900',     # index: 3,     light: -3,  CYMK:   0   25      40      0
                                      'FFFF00',     # index: 5,     light: -3,  CYMK:   0   0       40      0
                                      '00FF19',     # index: 7,     light: -3,  CYMK:   45  0       35      0
                                      '0019FF',     # index: 9,     light: -3,  CYMK:   50  25      0       0
                                      '9900FF',     # index: 11,    light: -3,  CYMK:   40  50      0       0

                                      'FF0000',     # index: 1,     light: -2,  CYMK:   0   65      50      0
                                      'FF5900',     # index: 3,     light: -2,  CYMK:   0   40     60       0
                                      'FFFF00',     # index: 5,     light: -2,  CYMK:   0   0       0       60
                                      '00FF19',     # index: 7,     light: -2,  CYMK:   60  0       55      0
                                      '0019FF',     # index: 9,     light: -2,  CYMK:   65  40      0       0
                                      '9900FF',     # index: 11,    light: -2,  CYMK:   55  65      0       0

                                      'FF0000',     # index: 1,     light: -4,  CYMK:   0   25      15      0
                                      'FF5900',     # index: 3,     light: -4,  CYMK:   0   15      25      0
                                      'FFFF00',     # index: 5,     light: -4,  CYMK:   0   0       20      0
                                      '00FF19',     # index: 7,     light: -4,  CYMK:   25  0       20      0
                                      '0019FF',     # index: 9,     light: -4,  CYMK:   30  15      0       0
                                      '9900FF',     # index: 11,    light: -4,  CYMK:   25  30     0        0

                                      'FF0000',     # index: 2,     light: -1,  CYMK:   0   70      65      0
                                      'FF5900',     # index: 4,     light: -1,  CYMK:   0   30      80      0
                                      'FFFF00',     # index: 6,     light: -1,  CYMK:   50  0       80      0
                                      '00FF19',     # index: 8,     light: -1,  CYMK:   80  0       30      0
                                      '0019FF',     # index: 10,    light: -1,  CYMK:   85  80      0       0
                                      '9900FF',     # index: 12,    light: -1,  CYMK:   35  80      0       0

                                      'FF0000',     # index: 2,     light: -3,  CYMK:   0   40      35      0
                                      'FF5900',     # index: 4,     light: -3,  CYMK:   0   15      40      0
                                      'FFFF00',     # index: 6,     light: -3,  CYMK:   25  0       40      0
                                      '00FF19',     # index: 8,     light: -3,  CYMK:   45  0       20      0
                                      '0019FF',     # index: 10,    light: -3,  CYMK:   60  55      0       0
                                      '9900FF',     # index: 12,    light: -3,  CYMK:   20  40      0       0

                                      'FF0000',     # index: 2,     light: -2,  CYMK:   0   55      50      0
                                      'FF5900',     # index: 4,     light: -2,  CYMK:   0   25      60      0
                                      'FFFF00',     # index: 6,     light: -2,  CYMK:   35  0       60      0
                                      '00FF19',     # index: 8,     light: -2,  CYMK:   60  0       25      0
                                      '0019FF',     # index: 10,    light: -2,  CYMK:   75  65      0       0
                                      '9900FF',     # index: 12,    light: -2,  CYMK:   25  60      0       0

                                      'FF0000',     # index: 2,     light: -4,  CYMK:   0   20      20      0
                                      'FF5900',     # index: 4,     light: -4,  CYMK:   0   10      20      0
                                      'FFFF00',     # index: 6,     light: -4,  CYMK:   10  0       20      0
                                      '00FF19',     # index: 8,     light: -4,  CYMK:   25  0       10      0
                                      '0019FF',     # index: 10,    light: -4,  CYMK:   45  40      0       0
                                      '9900FF'],    # index: 12,    light: -4,  CYMK:   10  20      0       0
        '''
        self.line_route_colors_in = [['ff0000',  # index: 1,     light: 0,   CYMK:   0   100     100     0
                                      'ff5900',  # index: 3,     light: 0,   CYMK:   0   65      100     0
                                      'ffff00',  # index: 5,     light: 0,   CYMK:   0   0       100     0
                                      '00ff19',  # index: 7,     light: 0,   CYMK:   100 0       90      0
                                      '0066ff',  # index: 9,     light: 0,   CYMK:   100 60      0       0
                                      '3200ff',  # index: 11,    light: 0,   CYMK:   80  100     0       0

                                      'b20000',  # index: 1,     light: 30,  CYMK:   0   100     100     30
                                      'b23e00',  # index: 3,     light: 30,  CYMK:   0   65      100     30
                                      'b2b200',  # index: 5,     light: 30,  CYMK:   0   0       100     30
                                      '00b211',  # index: 7,     light: 30,  CYMK:   100 0       90      30
                                      '0047b2',  # index: 9,     light: 30,  CYMK:   100 60      0       30
                                      '2300b2',  # index: 11,    light: 30,  CYMK:   80  100     0       30

                                      '660000',  # index: 1,     light: 60,  CYMK:   0   100     100     60
                                      '662300',  # index: 3,     light: 60,  CYMK:   0   65      100     60
                                      '666600',  # index: 5,     light: 60,  CYMK:   0   0       100     60
                                      '00660a',  # index: 7,     light: 60,  CYMK:   100 0       90      60
                                      '002866',  # index: 9,     light: 60,  CYMK:   100 60      0       60
                                      '140066',  # index: 11,    light: 60,  CYMK:   80  100     0       60

                                      'ff1932',  # index: 2,     light: 0,   CYMK:   0   90      80      0
                                      'ff9900',  # index: 4,     light: 0,   CYMK:   0   40      100     0
                                      '66ff00',  # index: 6,     light: 0,   CYMK:   60  0       100     0
                                      '00ff99',  # index: 8,     light: 0,   CYMK:   100 0       40      0
                                      '0019ff',  # index: 10,    light: 0,   CYMK:   100 90      0       0
                                      '9900ff',  # index: 12,    light: 0,   CYMK:   40  100     0       0

                                      'b21123',  # index: 2,     light: 30,  CYMK:   0   90      80      30
                                      'b26b00',  # index: 4,     light: 30,  CYMK:   0   40      100     30
                                      '47b200',  # index: 6,     light: 30,  CYMK:   60  0       100     30
                                      '00b26b',  # index: 8,     light: 30,  CYMK:   100 0       40      30
                                      '0011b2',  # index: 10,    light: 30,  CYMK:   100 90      0       30
                                      '6b00b2',  # index: 12,    light: 30,  CYMK:   40  100     0       30

                                      '660a14',  # index: 2,     light: 60,  CYMK:   0   90      80      60
                                      '663d00',  # index: 4,     light: 60,  CYMK:   0   40      100     60
                                      '286600',  # index: 6,     light: 60,  CYMK:   60  0       100     60
                                      '00663d',  # index: 8,     light: 60,  CYMK:   100 0       40      60
                                      '000a66',  # index: 10,    light: 60,  CYMK:   100 90      0       60
                                      '3d0066',  # index: 12,    light: 60,  CYMK:   40  100     0       60

                                      'd80000',  # index: 1,     light: 15,  CYMK:   0   100     100     15
                                      'd84b00',  # index: 3,     light: 15,  CYMK:   0   65      100     15
                                      'd8d800',  # index: 5,     light: 15,  CYMK:   0   0       100     15
                                      '00d815',  # index: 7,     light: 15,  CYMK:   100 0       90      15
                                      '0056d8',  # index: 9,     light: 15,  CYMK:   100 60      0       15
                                      '2300b2',  # index: 11,    light: 15,  CYMK:   80  100     0       15

                                      '8c0000',  # index: 1,     light: 45,  CYMK:   0   100     100     45
                                      '8c3100',  # index: 3,     light: 45,  CYMK:   0   65      100     45
                                      '8c8c00',  # index: 5,     light: 45,  CYMK:   0   0       100     45
                                      '008c0e',  # index: 7,     light: 45,  CYMK:   100 0       90      45
                                      '00388c',  # index: 9,     light: 45,  CYMK:   100 60      0       45
                                      '1c008c',  # index: 11,    light: 45,  CYMK:   80  100     0       45

                                      'd8152b',  # index: 2,     light: 15,  CYMK:   0   90      80      15
                                      'd88200',  # index: 4,     light: 15,  CYMK:   0   40      100     15
                                      '56d800',  # index: 6,     light: 15,  CYMK:   60  0       100     15
                                      '00d882',  # index: 8,     light: 15,  CYMK:   100 0       40      15
                                      '0015d8',  # index: 10,    light: 15,  CYMK:   100 90      0       15
                                      '8200d8',  # index: 12,    light: 15,  CYMK:   40  100     0       15

                                      '8c0e1c',  # index: 2,     light: 45,  CYMK:   0   90      80      45
                                      '8c5400',  # index: 4,     light: 45,  CYMK:   0   40      100     45
                                      '388c00',  # index: 6,     light: 45,  CYMK:   60  0       100     45
                                      '008c54',  # index: 8,     light: 45,  CYMK:   100 0       40      45
                                      '000e8c',  # index: 10,    light: 45,  CYMK:   100 90      0       45
                                      '54008c',  # index: 12,    light: 45,  CYMK:   40  100     0       45

                                      'ff264c',  # index: 1,     light: -1,  CYMK:   0   85      75      0
                                      'ff7f32',  # index: 3,     light: -1,  CYMK:   0   50      80      0
                                      '323232',  # index: 5,     light: -1,  CYMK:   0   0       0       80
                                      '32ff3f',  # index: 7,     light: -1,  CYMK:   0   100     100     0
                                      '267fff',  # index: 9,     light: -1,  CYMK:   85  50      0       0
                                      '5926ff',  # index: 11,    light: -1,  CYMK:   65  85      0       0

                                      'ff8cb2',  # index: 1,     light: -3,  CYMK:   0   45      30      0
                                      'ffbf99',  # index: 3,     light: -3,  CYMK:   0   25      40      0
                                      'ffff99',  # index: 5,     light: -3,  CYMK:   0   0       40      0
                                      '8cffa5',  # index: 7,     light: -3,  CYMK:   45  0       35      0
                                      '7fbfff',  # index: 9,     light: -3,  CYMK:   50  25      0       0
                                      '997fff',  # index: 11,    light: -3,  CYMK:   40  50      0       0

                                      'ff597f',  # index: 1,     light: -2,  CYMK:   0   65      50      0
                                      'ff9966',  # index: 3,     light: -2,  CYMK:   0   40     60       0
                                      '666666',  # index: 5,     light: -2,  CYMK:   0   0       0       60
                                      '66ff72',  # index: 7,     light: -2,  CYMK:   60  0       55      0
                                      '5999ff',  # index: 9,     light: -2,  CYMK:   65  40      0       0
                                      '7259ff',  # index: 11,    light: -2,  CYMK:   55  65      0       0

                                      'ffbfd8',  # index: 1,     light: -4,  CYMK:   0   25      15      0
                                      'ffd8cc',  # index: 3,     light: -4,  CYMK:   0   15      25      0
                                      'ffffcc',  # index: 5,     light: -4,  CYMK:   0   0       20      0
                                      'bfffcc',  # index: 7,     light: -4,  CYMK:   25  0       20      0
                                      'b2d8ff',  # index: 9,     light: -4,  CYMK:   30  15      0       0
                                      'bfb2ff',  # index: 11,    light: -4,  CYMK:   25  30     0        0

                                      'ff4c59',  # index: 2,     light: -1,  CYMK:   0   70      65      0
                                      'ffb232',  # index: 4,     light: -1,  CYMK:   0   30      80      0
                                      '7fff32',  # index: 6,     light: -1,  CYMK:   50  0       80      0
                                      '32ffb2',  # index: 8,     light: -1,  CYMK:   80  0       30      0
                                      '2632ff',  # index: 10,    light: -1,  CYMK:   85  80      0       0
                                      '3232ff',  # index: 12,    light: -1,  CYMK:   35  80      0       0

                                      'ff99a5',  # index: 2,     light: -3,  CYMK:   0   40      35      0
                                      'ffd899',  # index: 4,     light: -3,  CYMK:   0   15      40      0
                                      'bfff99',  # index: 6,     light: -3,  CYMK:   25  0       40      0
                                      '8cffcc',  # index: 8,     light: -3,  CYMK:   45  0       20      0
                                      '6672ff',  # index: 10,    light: -3,  CYMK:   60  55      0       0
                                      'cc99ff',  # index: 12,    light: -3,  CYMK:   20  40      0       0

                                      'ff727f',  # index: 2,     light: -2,  CYMK:   0   55      50      0
                                      'ffbf66',  # index: 4,     light: -2,  CYMK:   0   25      60      0
                                      'a5ff66',  # index: 6,     light: -2,  CYMK:   35  0       60      0
                                      '66ffbf',  # index: 8,     light: -2,  CYMK:   60  0       25      0
                                      '3f59ff',  # index: 10,    light: -2,  CYMK:   75  65      0       0
                                      'bf66ff',  # index: 12,    light: -2,  CYMK:   25  60      0       0

                                      'ffcccc',  # index: 2,     light: -4,  CYMK:   0   20      20      0
                                      'ffe5cc',  # index: 4,     light: -4,  CYMK:   0   10      20      0
                                      'e5ffcc',  # index: 6,     light: -4,  CYMK:   10  0       20      0
                                      'bfffe5',  # index: 8,     light: -4,  CYMK:   25  0       10      0
                                      '8c99ff',  # index: 10,    light: -4,  CYMK:   45  40      0       0
                                      'e5ccff'],  # index: 12,    light: -4,  CYMK:   10  20      0       0
                                     1,
                                     1,
                                     2.5]
        self.line_route_colors_in[0] = [''.join(["ff", color]) for color in
                                        self.line_route_colors_in[0]]  # ff for Visum usage
        self.line_route_colors_out = ['ff000000',
                                      1,
                                      1,
                                      3]


## set attributes
# define separator of folder in folder pathes
sep_folder = "\\"

folder_paths = VisumFolderStructure(Visum, sep_folder)
gpa = VisumLineRoute()
scenario_name = Visum.Net.AttValue("scenario_name")
lintim_version_name = Visum.Net.AttValue("lintim_version_name")

name_gpar = scenario_name + "_" + "Lines" + "_" + lintim_version_name

# write graphic parameter file
with open( folder_paths.graphicparameter + name_gpar + ".gpa", "w") as f:
    f.write(gpa.header)
    lines = Visum.Net.Lines.GetMultiAttValues("NAME")
    index = 0
    for line in lines:
        print(line[1] + "   " + str(index) + "   " + str(len(lines)))
        if index >= len(gpa.line_route_colors_in[0]) - 1:
            index = 0
        f.write(''.join([gpa.net_obj,
                         line[1] + '"' + gpa.closer + '\n',
                         gpa.line_style,
                         gpa.stokes,
                         ''.join([gpa.stoke_color,
                                  str(gpa.line_route_colors_out[0]) + '" ',
                                  gpa.dash_pattern,
                                  str(gpa.line_route_colors_out[1]) + '" ',
                                  gpa.dash_width,
                                  str(gpa.line_route_colors_out[2]) + '" ',
                                  gpa.width,
                                  str(gpa.line_route_colors_out[3]) + '"',
                                  gpa.closer_slash + '\n']),
                         ''.join([gpa.stoke_color,
                                  str(gpa.line_route_colors_in[0][index]) + '" ',
                                  gpa.dash_pattern,
                                  str(gpa.line_route_colors_in[1]) + '" ',
                                  gpa.dash_width,
                                  str(gpa.line_route_colors_in[2]) + '" ',
                                  gpa.width,
                                  str(gpa.line_route_colors_in[3]) + '"',
                                  gpa.closer_slash + '\n']),
                         gpa.stokes_closer,
                         gpa.line_style_closer,
                         gpa.net_obj_closer]))
        index += 1
    f.write(gpa.footer)

# Read graphic parameter file
Visum.Net.GraphicParameters.OpenXml(sep_folder.join([folder_paths.graphicparameter, name_gpar + ".gpa"]))

