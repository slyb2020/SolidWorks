import win32com.client as win32
import pythoncom
from swconst import constants

swApp = win32.Dispatch("Sldworks.application")            #引入sldworks接口
swApp.Visible = True                                      #是否可视化
arg_Nothing = win32.VARIANT(pythoncom.VT_DISPATCH, None)   #转义VBA中不同变量nothing


# os.system("E:\\SOLIDWORKS Corp\\SOLIDWORKS\\SLDWORKS.exe")#运行SolidWorks
# swApp.NewDocument("C:\\ProgramData\\SolidWorks\\SOLIDWORKS 2019\\templates\\gb_part.prtdot", 0, 0, 0)#新建一个国标零件
class StraightWallPanel(object):
    def __init__(self):
        super(StraightWallPanel, self).__init__()
        self.MakeXSurface(K=0.5,R=0.001,thickness=0.00075,height=1.6,width=0.55,bendLeft=0.015,bendRight=0.015)


    def MakeXSurface(self,K,R,thickness,height,width,bendLeft,bendRight):
        Part = swApp.ActiveDoc
        boolstatus = Part.Extension.SelectByID2("上视基准面", "PLANE", 1, 1, 1, False, 0, arg_Nothing, 0)#选择作图的基准面，这句很如果没有画出来的图方向就不一定正确，应该是使用二零默认的基准面
        Part.SketchManager.InsertSketch(True)#这句话必须有，代表新建一个草图
        skSegment = Part.SketchManager.CreateLine(0, 0, 0, 0, bendLeft, 0)#对于上视基准面，坐标顺序是x,z,y,单位是米
        skSegment = Part.SketchManager.CreateLine(0, 0, 0, width, 0, 0)
        skSegment = Part.SketchManager.CreateLine(width, 0, 0, width, bendRight, 0)
        # Part.ClearSelection2(True)

        customBendAllowanceData = Part.FeatureManager.CreateCustomBendAllowance   #生成钣金参数对象
        customBendAllowanceData.KFactor = K   #设定K系数
        myFeature = Part.FeatureManager.InsertSheetMetalBaseFlange2(thickness, False, R, height, R, False, 0, 0, 1, customBendAllowanceData, False, 0, 0.0001, 0.0001, 1, True, False, True, True)
        # Part.ShowNamedView2("*上下二等角轴测", 7)
        # Part.ViewZoomtofit2

if __name__=="__main__":
    N2AS255 = StraightWallPanel()