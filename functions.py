import os 
import sys
import comtypes.client 


def connect_to_etabs():
    """
    
    Attaching to a Manually Started Instance of ETABS 
    
    Returns:
    SapModel: type cOAPI pointer
    """

    # Create API helper object
    helper = comtypes.client.CreateObject('ETABSv1.Helper')
    helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
    
    # Attach to a running instance of ETABS
    try:
        # Get the active ETABS object
        myETABSObject = helper.GetObject("CSI.ETABS.API.ETABSObject") 
    except (OSError, comtypes.COMError):
        print("No running instance of the program found or failed to attach.")
        sys.exit(-1)
    
    # Create SapModel object
    SapModel = myETABSObject.SapModel
    
    return SapModel


def get_story_elevations(SapModel):
    """
    Retrieves the story elevations of a tower in the model.
    
    Parameters:
        :param SapModel: cOAPI indicator
    
    Returns:
    a list of the retrieved story elevations
    
    story_elevations: list | float | int 
    """
    
    # Get the data using API
    story_info = SapModel.Story.GetStories_2()
    
    # Extract the story elevations into a list
    story_elevations = list(story_info[3])
    
    return story_elevations;


def get_combinations(SapModel):
    """
    Retrieves the available load Combinations into a list.
    
    Parameters:
        :param SapModel: cOAPI indicator
        
    Returns:
    A list of load Combinations
    
    combos : list
    """
    
    combos = list(SapModel.RespCombo.GetNameList()[1])
    
    return combos;


def jointDisp_export(SapModel, savepath):
    """
    Exports a dataset of joint displacements from ETABS into a spreadsheet file in a desired path.
    
    Parameters:
        :param SapModel: cOAPI indicator
        :param savepath | string, the specified path as to which the spreadsheet file is to be saved
    
    Returns:
    Returns spreadsheet file to the specified directory
    
    Exception:
    Function raises an exception when 'Joint Displacement' table is absent in ETABS, ensure that analysis has been run.
    
    Examples:
    >> jointDisp_export(SapModel, savepath = 'disp-values.csv')
    Verify file existence at desktop, ✅
    
    >>  jointDisp_export(SapModel, savepath = 'disp-values.xlsx')
    Verify file existence at desktop, ✅
    """
    
    # Retrieve the table key with API (i.e 'Joint Displacement')
    key = SapModel.DatabaseTables.GetAvailableTables()[2][62]
    
    # Retrieve all the fields in the table (e.g ['Story', 'Label', 'Unique Name'..])
    fieldList = list(SapModel.DatabaseTables.GetAllFieldsInTable(TableKey = ti)[2])
    
    # OPTIONAL unless you want these fields removed.
    # unwanted_fields = ['StepType', 'StepNumber', 'StepLabel', 'Rx', 'Ry', 'Rz', 'Ux', 'Uy']
    # fieldList = [field for field in fieldList if field not in unwanted_fields]
    
    # Export the spreadsheet file to the desired path
    SapModel.DatabaseTables.GetTableForDisplayCSVFile(TableKey = key,
                                                      FieldKeyList = fieldList,
                                                      GroupName = 'All',
                                                      csvFilePath = savepath);
def columns(SapModel):
    """
    Returns a list of all column unique names from the etabs model,
    and all column end points of the Etabs model
    
    Parameters:
    -----------
        :param SapModel: cOAPI indicator
    
    Returns:
    out: columns: list
         joints: set
    """
    
    # Extract all the frame unique names from Etabs
    frames = SapModel.FrameObj.GetAllFrames()[1]
    
    # Extract all the frame end points from Etabs
    pointName1 = SapModel.FrameObj.GetAllFrames()[3]
    pointName2 = SapModel.FrameObj.GetAllFrames()[4]
    
    # Extract all the frame joint geometry from Etabs
    point1X = SapModel.FrameObj.GetAllFrames()[6]
    point1Y = SapModel.FrameObj.GetAllFrames()[7]
    point1Z = SapModel.FrameObj.GetAllFrames()[8]
    
    point2X = SapModel.FrameObj.GetAllFrames()[9]
    point2Y = SapModel.FrameObj.GetAllFrames()[10]
    point2Z = SapModel.FrameObj.GetAllFrames()[11]
    
    # Append all vertical columns into a list, 
    # Frame is a vertical column if x and y coordinates of the end points are the same
    # Append the joints of these columns into a list as well
    columns = []
    joints = set()
    for eachframe in frames:
        i = frames.index(eachframe)
        j = SapModel.FrameObj.GetPoints(eachframe)[0:2]
        
        if [point1X[i], point1Y[i]] == [point2X[i], point2Y[i]]:
            columns.append(eachframe)
            joints.update(j)
    
    return (columns, joints)
