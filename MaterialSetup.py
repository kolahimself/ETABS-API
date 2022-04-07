"""
Author: James Kola Ojoawo
Date: 8th March, 2022   Time: 8:14pm

This script uses the ETABS API to create Material Properties in a corresponding ETABS version.
Future commitments include a GUI & faster optimization in the software itself. 
"""

# Perform Necessary Imports
import os
import sys
import comtypes.client

# Print program purpose
purpose = """
The purpose of this program is simply an automation tool,
this program creates a material type property in ETABS which can be further modified.
"""
print(purpose)

# set the following flag to True to attach to an existing instance of the program
# otherwise a new instance of the program will be started
instance = int(input('Input 1 to attach to an existing instance of the program: '
                     'Input 0 to start new instance of the program: '))
if instance == 1:
    AttachToInstance = True
else:
    AttachToInstance = False

# this allows for a connection to a version of ETABS other than the latest installation
ExeDirectory = int(input('Input 1 if you want a connection to a version of ETABS other than the latest installation:\n '
                         'Input 0 if otherwise(latest installed version of ETABS will be launched): '))
if ExeDirectory == 1:
    SpecifyPath = True
    ProgramPath = str(input('Specify the path to ETABS: '))
else:
    SpecifyPath = False

# full path to the model
# set it to the desired path of your model
APIPath = str(input('Specify the desired part of your ETABS API Model: '))
if not os.path.exists(APIPath):
    try:
        os.makedirs(APIPath)
    except OSError:
        pass
ModelPath = APIPath + os.sep + 'API_1-001.edb'

# create API helper object
helper = comtypes.client.CreateObject('ETABSv1.Helper')
helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)

if AttachToInstance:
    # attach to a running instance of ETABS
    try:
        # get the active ETABS object
        myETABSObject = helper.GetObject("CSI.ETABS.API.ETABSObject")
    except (OSError, comtypes.COMError):
        print("No running instance of the program found or failed to attach.")
        sys.exit(-1)
else:
    if SpecifyPath:
        try:
            # 'create an instance of the ETABS object from the specified path
            myETABSObject = helper.CreateObject(ProgramPath)
        except (OSError, comtypes.COMError):
            print("Cannot start a new instance of the program from " + ProgramPath)
            sys.exit(-1)
    else:
        try:
            # create an instance of the ETABS object from the latest installed ETABS
            myETABSObject = helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject")
        except (OSError, comtypes.COMError):
            print("Cannot start a new instance of the program.")
            sys.exit(-1)

    # start ETABS application
    myETABSObject.ApplicationStart()

# create SapModel object
SapModel = myETABSObject.SapModel

# initialize model
SapModel.InitializeNewModel()

# create new blank model
ret = SapModel.File.NewBlank()

# Accept User Material Preferences
name = str(input('Supply the desired material name you want to create: '))
region = int(input('The following regions are available in ETABS: '
                   '0.  China'
                   '1.  Europe'
                   '2.  India'
                   '3.  Italy'
                   '4.  Korea'
                   '5.  New Zealand'
                   '6.  Russia'
                   '7.  Spain'
                   '8.  United States'
                   '9.  Vietnam'
                   '10. User'))

members = int(input('The following members are available in ETABS: '
                    '0.  Steel'
                    '1.  Concrete'
                    '2.  NoDesign'
                    '3.  Aluminium'
                    '4.  ColdFormed'
                    '5.  Rebar'
                    '6.  Tendon'
                    '7.  Masonry'
                    '8.  Other'))

members_list = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
region_list = ['China', 'Europe', 'India', 'Italy', 'Korea', 'New Zealand', 'Russia',
               'Spain', 'United States', 'Vietnam', 'User']

# define material property
ret = SapModel.PropMaterial.AddMaterial(name, members_dict[members], region_list[region])
