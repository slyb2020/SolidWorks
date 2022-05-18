import win32com.client as win32
import pythoncom
from swconst import constants

swApp = win32.Dispatch("Sldworks.application")            #引入sldworks接口
swApp.Visible = True                                      #是否可视化
arg_Nothing = win32.VARIANT(pythoncom.VT_DISPATCH, None)   #转义VBA中不同变量nothing


# os.system("E:\\SOLIDWORKS Corp\\SOLIDWORKS\\SLDWORKS.exe")#运行SolidWorks
# swApp.NewDocument("C:\\ProgramData\\SolidWorks\\SOLIDWORKS 2019\\templates\\gb_part.prtdot", 0, 0, 0)#新建一个国标零件

Part = swApp.ActiveDoc
boolstatus = Part.Extension.SelectByID2("上视基准面", "PLANE", 1, 1, 1, False, 0, arg_Nothing, 0)
Part.SketchManager.InsertSketch(True)
Part.ClearSelection2(True)
skSegment = Part.SketchManager.CreateLine(0, 0, 0, 0, 0.007, 0)#坐标顺序是x,z,y,单位是米
skSegment = Part.SketchManager.CreateLine(0, 0, 0, 0.55, 0, 0)
skSegment = Part.SketchManager.CreateLine(0.55, 0, 0, 0.55, 0.007, 0)
Part.ClearSelection2(True)

# ' Named View
Part.ShowNamedView2("*上下二等角轴测", 8)
Part.ViewZoomtofit2
customBendAllowanceData = Part.FeatureManager.CreateCustomBendAllowance
customBendAllowanceData.KFactor = 0.5
myFeature = Part.FeatureManager.InsertSheetMetalBaseFlange2(0.00075, False, 0.001, 1.6, 0.01, False, 0, 0, 1, customBendAllowanceData, False, 0, 0.0001, 0.0001, 1, True, False, True, True)
Part.ShowNamedView2("*上下二等角轴测", 7)
