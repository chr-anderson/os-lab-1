import clr, os, winreg
from itertools import islice

# determine the Zemax working directory
aKey = winreg.OpenKey(winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER), r"Software\Zemax", 0, winreg.KEY_READ)
zemaxData = winreg.QueryValueEx(aKey, 'ZemaxRoot')
NetHelper = os.path.join(os.sep, zemaxData[0], r'ZOS-API\Libraries\ZOSAPI_NetHelper.dll')
winreg.CloseKey(aKey)

# add the NetHelper DLL for locating the OpticStudio install folder
clr.AddReference(NetHelper)
import ZOSAPI_NetHelper

pathToInstall = ''
# uncomment the following line to use a specific instance of the ZOS-API assemblies
#pathToInstall = r'C:\C:\Program Files\Zemax OpticStudio'

# connect to OpticStudio
success = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(pathToInstall);

zemaxDir = ''
if success:
    zemaxDir = ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory();
    print('Found OpticStudio at:   %s' + zemaxDir);
else:
    raise Exception('Cannot find OpticStudio')

# load the ZOS-API assemblies
clr.AddReference(os.path.join(os.sep, zemaxDir, r'ZOSAPI.dll'))
clr.AddReference(os.path.join(os.sep, zemaxDir, r'ZOSAPI_Interfaces.dll'))
import ZOSAPI

TheConnection = ZOSAPI.ZOSAPI_Connection()
if TheConnection is None:
    raise Exception("Unable to intialize NET connection to ZOSAPI")

TheApplication = TheConnection.ConnectAsExtension(0)
if TheApplication is None:
    raise Exception("Unable to acquire ZOSAPI application")

if TheApplication.IsValidLicenseForAPI == False:
    raise Exception("License is not valid for ZOSAPI use.  Make sure you have enabled 'Programming > Interactive Extension' from the OpticStudio GUI.")

TheSystem = TheApplication.PrimarySystem
if TheSystem is None:
    raise Exception("Unable to acquire Primary system")

def reshape(self, data, x, y, transpose = False):
    """Converts a System.Double[,] to a 2D list for plotting or post processing

    Parameters
    ----------
    data      : System.Double[,] data directly from ZOS-API
    x         : x width of new 2D list [use var.GetLength(0) for dimension]
    y         : y width of new 2D list [use var.GetLength(1) for dimension]
    transpose : transposes data; needed for some multi-dimensional line series data

    Returns
    -------
    res       : 2D list; can be directly used with Matplotlib or converted to
                a numpy array using numpy.asarray(res)
    """
    if type(data) is not list:
        data = list(data)
    var_lst = [y] * x;
    it = iter(data)
    res = [list(islice(it, i)) for i in var_lst]
    if transpose:
        return self.transpose(res);
    return res

def transpose(self, data):
    """Transposes a 2D list (Python3.x or greater).

    Useful for converting mutli-dimensional line series (i.e. FFT PSF)

    Parameters
    ----------
    data      : Python native list (if using System.Data[,] object reshape first)

    Returns
    -------
    res       : transposed 2D list
    """
    if type(data) is not list:
        data = list(data)
    return list(map(list, zip(*data)))

print('Connected to OpticStudio')

# The connection should now be ready to use.  For example:
print('Serial #: ', TheApplication.SerialCode)

# Start and save a new project file at the specified filepath
TheSystem.New(False)
fileOut = '<save directory here>/APISinglet.zmx'
TheSystem.SaveAs(fileOut)

# Accessing system explorer data
print("Setting up in system explorer...")
TheSystemData = TheSystem.SystemData

#Setting entrance pupil diameter to 65 mm
TheSystemData.Aperture.ApertureValue = 65

# Setting wavelength to the HeNe 633 nm preset
TheSystemData.Wavelengths.SelectWavelengthPreset(ZOSAPI.SystemData.WavelengthPreset.HeNe_0p6328)
print("System set!")

# Entering lens data
print("Entering lens data...")
TheLDE = TheSystem.LDE
TheLDE.InsertNewSurfaceAt(1)
TheLDE.InsertNewSurfaceAt(1)

Surface_1 = TheLDE.GetSurfaceAt(1)
Surface_2 = TheLDE.GetSurfaceAt(2)
Surface_3 = TheLDE.GetSurfaceAt(3)

# Front of lens
Surface_1.Thickness = 15.0
Surface_1.Comment = "Front of lens"
Surface_1.Material = "N-BK7"

# Back of lens
Surface_2.Thickness = 60.0
Surface_2.Comment = "Back of lens"

# Stop
Surface_3.Comment = "Stop"
Surface_3.Thickness = 400.0

# Setting solve variables
Surface_1.RadiusCell.MakeSolveVariable()
Surface_1.ThicknessCell.MakeSolveVariable()
Surface_2.RadiusCell.MakeSolveVariable()
Surface_2.ThicknessCell.MakeSolveVariable()
Surface_3.ThicknessCell.MakeSolveVariable()
print("Lenses set!")

# Setting up the merit function
print("Setting up the merit function...")
TheMFE = TheSystem.MFE
wizard = TheMFE.SEQOptimizationWizard

wizard.Type = 0      # RMS
wizard.Data = 1      # Spot radius
wizard.Reference = 0 # Centroid
wizard.Ring = 2      # 3 rings
wizard.arm = 0       # 6 arms

# Boundary values for glass and air
wizard.IsGlassUsed = True
wizard.GlassMin = 3        # Min thickness
wizard.GlassMax = 15       # Max thickness
wizard.GlassEdge = 3       # Edge thickness

wizard.IsAirUsed = True
wizard.AirMin = 0.5       # Min thickness
wizard.AirMax = 1000      # Max thickness
wizard.AirEdge = 0.5      # Edge thickness

wizard.IsAssumeAxialSymmetryUsed = True
wizard.CommonSettings.OK()

# Stop is at the back of the system, so we use an effective focal length solve
EFLSolve = TheMFE.InsertNewOperandAt(1)
EFLSolve.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.EFFL)
EFLSolve.Target = 400.0
EFLSolve.Weight = 1.0
print("Merit function set!")

# Optimizing
print("Optimizing...")
LocalOpt = TheSystem.Tools.OpenLocalOptimization()
LocalOpt.Algorithm = ZOSAPI.Tools.Optimization.OptimizationAlgorithm.DampedLeastSquares
LocalOpt.Cycles = ZOSAPI.Tools.Optimization.OptimizationCycles.Automatic
LocalOpt.NumberOfCores = 8
LocalOpt.RunAndWaitForCompletion()
LocalOpt.Close()
print("Optimized!")

# Save and disconnect
TheSystem.Save()
TheApplication.CloseApplication()
