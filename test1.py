import win32com.client as win32
import pythoncom
from swconst import constants

swApp = win32.Dispatch("Sldworks.application")            #引入sldworks接口
swApp.Visible = True                                      #是否可视化
arg_Nothing = win32.VARIANT(pythoncom.VT_DISPATCH, None)   #转义VBA中不同变量nothing

# os.system("C:\\SOLIDWORKS Corp\\SOLIDWORKS\\SLDWORKS.exe")#运行SolidWorks
swApp.NewDocument("C:\\ProgramData\\SolidWorks\\SOLIDWORKS 2019\\templates\\gb_part.prtdot", 0, 0, 0)#新建一个国标零件
# longstatus = Part.SaveAs3("C:\\Users\\xls\\Desktop\\sw_exercise\\test1.SLDPRT", 0, 2)
# swApp.ActivateDoc2("零件1", False, longstatus)
# myModelView = Part.ActiveView  #这两句是试图将SolidWorks窗口最大化，没成功，还需进一步研究
# myModelView.FrameState = swWindowState_e.swWindowMaximized
Part = swApp.ActiveDoc #选择当前激活的零件图
Part.SketchManager.InsertSketch(True) #新建一个草图
skSegment1 = Part.SketchManager.CreateCircle(0, 0, 0, 1, 0, 0)
Part.SketchManager.InsertSketch(True)
myFeature1 = Part.FeatureManager.FeatureExtrusion2(True,False,False,0,0,0.5,0.5,False,False,False,False,0.1,0.1,False,False,False,False,True,True,True,0,0,False)



# skSegment1 = Part.SketchManager.CreateLine(0.16594, 0.031426, 0, 0.16594, 0.312677, 0)



# skSegment1 = Part.SketchManager.CreateLine(0.16594, 0.031426, 0, 0.16594, 0.312677, 0)
# skSegment2 = Part.SketchManager.CreateLine(0.16594, 0.312677, 0, 0.203282, 0.312677, 0)
# skSegment3 = Part.SketchManager.CreateLine(0.203282, 0.312677, 0, 0.203282, 0.286459, 0)
# skSegment4 = Part.SketchManager.CreateLine(0.203282, 0.286459, 0, 0.235061, 0.286459, 0)
# skSegment5 = Part.SketchManager.CreateLine(0.235061, 0.286459, 0, 0.235061, 0.261035, 0)
# skSegment6 = Part.SketchManager.CreateLine(0.235061, 0.261035, 0, 0.267636, 0.261035, 0)
# skSegment7 = Part.SketchManager.CreateLine(0.267636, 0.261035, 0, 0.267636, -0.044846, 0)
# skSegment8 = Part.SketchManager.CreateLine(0.267636, -0.044846, 0, 0.219171, -0.044846, 0)
# skSegment9 = Part.SketchManager.CreateLine(0.219171, -0.044846, 0, 0.219171, 0.031426, 0)
# skSegment10 = Part.SketchManager.CreateLine(0.219171, 0.031426, 0, 0.16594, 0.031426, 0)
#
# boolstatus1 = Part.Extension.SelectByID2("Line20", "SKETCHSEGMENT", 0.201692557177639, 3.22201012003451E-02, 0, False, 0, arg_Nothing, 0)
# myDisplayDim1 = Part.AddDimension2(0.198514570947221, -9.09268652283702E-02, 0)
# myDimension1 = Part.Parameter("D1@草图1")
# #myDimension1.SystemValue = 0.02
# boolstatus2 = Part.Extension.SelectByID2("Line18", "SKETCHSEGMENT", 0.22076047456015, -4.72295545601163E-02, 0, False, 0, arg_Nothing, 0)
# myDisplayDim2 = Part.AddDimension2(0.223938460790569, -0.129062699993392, 0)
# myDimension2 = Part.Parameter("D2@草图1")
# #myDimension2.SystemValue = 0.028
# boolstatus3 = Part.Extension.SelectByID2("Line16", "SKETCHSEGMENT", 0.214298450704151, 0.260911495250115, 0, False, 0, arg_Nothing, 0)
# myDisplayDim3 = Part.AddDimension2(0.214770455709364, 0.258023935218222, 0)
# myDimension3 = Part.Parameter("D3@草图1")
# #myDimension3.SystemValue = -0.014
# boolstatus4 = Part.Extension.SelectByID2("Line12", "SKETCHSEGMENT", 0.179105082440666, 0.312729478715755, 0, False, 0, arg_Nothing, 0)
# myDisplayDim4 = Part.AddDimension2(0.182612119139466, 0.326345032958157, 0)
# myDimension4 = Part.Parameter("D4@草图1")
# #myDimension4.SystemValue = 0.018
# boolstatus5 = Part.Extension.SelectByID2("Line13", "SKETCHSEGMENT", 0.184221550762286, 0.303894387116026, 0, False, 0, arg_Nothing, 0)
# myDisplayDim5 = Part.AddDimension2(0.224097644655966, 0.297248371467079, 0)
# myDimension5 = Part.Parameter("D5@草图1")
# #myDimension5.SystemValue = 0.012
# boolstatus6 = Part.Extension.SelectByID2("Line15", "SKETCHSEGMENT", 0.199930315023433, 0.287380045200461, 0, False, 0, arg_Nothing, 0)
# myDisplayDim6 = Part.AddDimension2(0.237993859194672, 0.281741001619537, 0)
# myDimension6 = Part.Parameter("D6@草图1")
# #myDimension6.SystemValue = 0.007
# boolstatus7 = Part.Extension.SelectByID2("Line11", "SKETCHSEGMENT", 0.16659311942184, 6.65606406038094E-02, 0, False, 0, arg_Nothing, 0)
# myDisplayDim7 = Part.AddDimension2(0.111708344759752, 6.60791952120367E-02, 0)
# myDimension7 = Part.Parameter("D7@草图1")
# #myDimension7.SystemValue = 0.273
# boolstatus8 = Part.Extension.SelectByID2("Line19", "SKETCHSEGMENT", 0.185850935092748, -0.044171799503913, 0, False, 0, arg_Nothing, 0)
# myDisplayDim8 = Part.AddDimension2(0.129040378863569, -5.04305895969581E-02, 0)
# myDimension8 = Part.Parameter("D8@草图1")
# #myDimension8.SystemValue = 0.037
# boolstatus9 = Part.Extension.SelectByID2("草图1", "SKETCH", 0, 0, 0, False, 4, arg_Nothing, 0)
# myFeature = Part.FeatureManager.FeatureExtrusion2(True, False, False, 6, 0, 2.2, 0.01, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, True, True, True, 0, 0, False)
# Part.SelectionManager.EnableContourSelection = False
#
# skSegment11 = Part.SketchManager.CreateLine(0.007915, 0.050299, 0, 0.038824, 0.050299, 0)
# skSegment12 = Part.SketchManager.CreateLine(0.038824, 0.050299, 0, 0.038824, -0.221772, 0)
# skSegment13 = Part.SketchManager.CreateLine(0.038824, -0.221772, 0, 0.007915, -0.221772, 0)
# skSegment14 = Part.SketchManager.CreateLine(0.007915, -0.221772, 0, 0.007915, -0.211662, 0)
# skSegment15 = Part.SketchManager.CreateLine(0.007915, -0.211662, 0, 0.028046, -0.211662, 0)
# skSegment16 = Part.SketchManager.CreateLine(0.028046, -0.211662, 0, 0.028046, 0.037169, 0)
# skSegment17 = Part.SketchManager.CreateLine(0.028046, 0.037169, 0, 0.007915, 0.037169, 0)
# skSegment18 = Part.SketchManager.CreateLine(0.007915, 0.037169, 0, 0.007915, 0.050299, 0)
#
# boolstatus10 = Part.Extension.SelectByID2("Line8", "SKETCHSEGMENT", 7.50599360498305E-03, 4.37260357277246E-02, 0, False, 0, arg_Nothing, 0)
# myDisplayDim10 = Part.AddDimension2(-2.15154013649017E-02, 4.17912760630656E-02, 0)
# myDimension10 = Part.Parameter("D1@草图1")
# #myDimension10.SystemValue = 0.0007
# boolstatus11 = Part.Extension.SelectByID2("Line4", "SKETCHSEGMENT", 3.52633883325251E-04, -1.15074609448399E-02, 0, False, 0,  arg_Nothing, 0)
# myDisplayDim11 = Part.AddDimension2(-6.4633309875551E-04, -1.15407598442426E-02, 0)
# myDimension11 = Part.Parameter("D2@草图1")
# #myDimension11.SystemValue = 0.0007
# boolstatus12 = Part.Extension.SelectByID2("Point7", "SKETCHPOINT", 1.49528678666582E-03, 1.98167562600092E-03, 0, False, 0, arg_Nothing, 0)
# boolstatus13 = Part.Extension.SelectByID2("Point2", "SKETCHPOINT", 2.0699075687632E-03, 2.68167562600092E-03, 0, True, 0,  arg_Nothing, 0)
# myDisplayDim13 = Part.AddDimension2(1.8656100029821E-03, 3.88141845971882E-03, 0)
# myDimension13 = Part.Parameter("D3@草图1")
# #myDimension13.SystemValue = 0.0007
# boolstatus14 = Part.Extension.SelectByID2("Line6", "SKETCHSEGMENT", 1.45908195017657E-03, 7.34091847882756E-04, 0, False, 0,  arg_Nothing, 0)
# myDisplayDim14 = Part.AddDimension2(-6.12663226153821E-04, -1.24949395924209E-03, 0)
# myDimension14 = Part.Parameter("D4@草图1")
# #myDimension14.SystemValue = 0.291
# boolstatus15 = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 1.19520079899023E-03, 0.280435486495549, 0, False, 0,  arg_Nothing, 0)
# myDisplayDim15 = Part.AddDimension2(1.36620921312163E-03, 0.282791602423582, 0)
# myDimension15 = Part.Parameter("D5@草图1")
# #myDimension15.SystemValue = 0.007
# boolstatus16 = Part.Extension.SelectByID2("Line3", "SKETCHSEGMENT", 1.29662075114369E-03, -1.20133375508271E-02, 0, False, 0,  arg_Nothing, 0)
# myDisplayDim16 = Part.AddDimension2(1.12757561967853E-03, -1.42743161841736E-02, 0)
# myDimension16 = Part.Parameter("D6@草图1")
# #myDimension16.SystemValue = 0.007
# boolstatus17 = Part.Extension.SelectByID2("Line3", "SKETCHSEGMENT", 5.88278665275517E-03, -1.20368171229151E-02, 0, True, 0, arg_Nothing, 0)
# boolstatus18 = Part.Extension.SelectByID2("圆角<1>", "SKETCHSEGMENT", 7.51322622003563E-03, -1.09992646710093E-02, 0, True, 0, arg_Nothing, 0)
# skSegment19 = Part.SketchManager.CreateFillet(0.0021, 1)
# boolstatus19 = Part.Extension.SelectByID2("Line5", "SKETCHSEGMENT", 4.86993783065671E-03, -1.14192263777331E-02, 0, True, 0, arg_Nothing, 0)
# boolstatus20 = Part.Extension.SelectByID2("圆角<1>", "SKETCHSEGMENT", 6.67330280658812E-03, -8.92415976719785E-03, 0, True, 0,  arg_Nothing, 0)
# skSegment20 = Part.SketchManager.CreateFillet(0.0014, 1)
# boolstatus21 = Part.Extension.SelectByID2("Line7", "SKETCHSEGMENT", 5.9063734381915E-03, 0.279709981460959, 0, True, 0,  arg_Nothing, 0)
# boolstatus22 = Part.Extension.SelectByID2("圆角<1>", "SKETCHSEGMENT", 6.77956818580351E-03, 0.278804840564044, 0, True, 0,  arg_Nothing, 0)
# skSegment21 = Part.SketchManager.CreateFillet(0.0014, 1)
# boolstatus23 = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 5.81053499028287E-03, 0.280412796745623, 0, True, 0,  arg_Nothing, 0)
# boolstatus24 = Part.Extension.SelectByID2("圆角<1>", "SKETCHSEGMENT", 7.40784245542679E-03, 0.278815489280479, 0, True, 0,  arg_Nothing, 0)
# skSegment22 = Part.SketchManager.CreateFillet(0.0021, 1)
# boolstatus25 = Part.Extension.SelectByID2("草图1", "SKETCH", 0, 0, 0, False, 4, arg_Nothing, 0)
# myFeature = Part.FeatureManager.FeatureExtrusion2(True, False, False, 6, 0, 2.2, 0.01, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, True, True, True, 0, 0, False)
# Part.SelectionManager.EnableContourSelection = False
#
# skSegment23 = Part.SketchManager.CreateLine(0.328091, -0.252698, 0, 0.370339, -0.229805, 0)
# skSegment24 = Part.SketchManager.CreateLine(0.370339, -0.229805, 0, 0.288148, -0.229805, 0)
# skSegment25 = Part.SketchManager.CreateLine(0.288148, -0.229805, 0, 0.288148, -0.159136, 0)
# skSegment26 = Part.SketchManager.CreateLine(0.288148, -0.159136, 0, 0.237451, -0.159136, 0)
# skSegment27 = Part.SketchManager.CreateLine(0.237451, -0.159136, 0, 0.237451, 0.08206, 0)
# skSegment28 = Part.SketchManager.CreateLine(0.237451, 0.08206, 0, 0.275089, 0.08206, 0)
# skSegment29 = Part.SketchManager.CreateLine(0.275089, 0.08206, 0, 0.275089, 0.060234, 0)
# skSegment30 = Part.SketchManager.CreateLine(0.275089, 0.060234, 0, 0.308888, 0.060234, 0)
# skSegment31 = Part.SketchManager.CreateLine(0.308888, 0.060234, 0, 0.291989, 0.043335, 0)
# skSegment32 = Part.SketchManager.CreateLine(0.291989, 0.043335, 0, 0.29529, 0.040034, 0)
# skSegment33 = Part.SketchManager.CreateLine(0.29529, 0.040034, 0, 0.31898, 0.063725, 0)
# skSegment34 = Part.SketchManager.CreateLine(0.31898, 0.063725, 0, 0.279803, 0.063725, 0)
# skSegment35 = Part.SketchManager.CreateLine(0.279803, 0.063725, 0, 0.279803, 0.089319, 0)
# skSegment36 = Part.SketchManager.CreateLine(0.279803, 0.089319, 0, 0.231836, 0.089319, 0)
# skSegment37 = Part.SketchManager.CreateLine(0.231836, 0.089319, 0, 0.231836, -0.164154, 0)
# skSegment38 = Part.SketchManager.CreateLine(0.231836, -0.164154, 0, 0.279803, -0.164154, 0)
# skSegment39 = Part.SketchManager.CreateLine(0.279803, -0.164154, 0, 0.279803, -0.235116, 0)
# skSegment40 = Part.SketchManager.CreateLine(0.279803, -0.235116, 0, 0.349215, -0.235116, 0)
# skSegment41 = Part.SketchManager.CreateLine(0.349215, -0.235116, 0, 0.325521, -0.247955, 0)
# skSegment42 = Part.SketchManager.CreateLine(0.325521, -0.247955, 0, 0.328091, -0.252698, 0)
