import win32com.client as wc

def combiner(h):
    oApp = wc.GetActiveObject('Inventor.Application')

    oPartDoc = oApp.Documents.Add(12290, oApp.FileManager.GetTemplateFile(12290, 8962))

    oSketch = oPartDoc.ComponentDefinition.Sketches.Add(oPartDoc.ComponentDefinition.WorkPlanes.Item(3))

    oTG = oApp.TransientGeometry

    #lower profile
    P1 = oTG.CreatePoint2d(0, 0.3)
    P2 = oTG.CreatePoint2d(0.0879 , 0.0879)
    P3 = oTG.CreatePoint2d(0.3, 0)

    P4 = oTG.CreatePoint2d(1.9121, -0.0458)
    P5 = oTG.CreatePoint2d(2.2577, -0.1806)
    P6 = oTG.CreatePoint2d(1.544, 0)

    P7 = oTG.CreatePoint2d(4.3794, -.7908)
    P8 = oTG.CreatePoint2d(2.9869, -.5751)
    P9 = oTG.CreatePoint2d(5.5947, -0.0778)

    P10 = oTG.CreatePoint2d(10.0204 , 4.2073)
    P11 = oTG.CreatePoint2d(8.5929, 3.4955)
    P12 = oTG.CreatePoint2d(11.5145, 3.6485)

    P13 = oTG.CreatePoint2d(15.3653, 0.182)
    P14 = oTG.CreatePoint2d(14.5651, .7018)
    P15 = oTG.CreatePoint2d(16.3020, 0)

    P16 = oTG.CreatePoint2d(18.4121, 0.0879)
    P17 = oTG.CreatePoint2d(18.2, 0)
    P18 = oTG.CreatePoint2d(18.5, .3)

    #upper profile
    P19 = oTG.CreatePoint2d(0.0879, 2.3221+h)
    P20 = oTG.CreatePoint2d(.3, 2.41+h)
    P21 = oTG.CreatePoint2d(0, 2.1100+h)

    P22 = oTG.CreatePoint2d(1.8856, 2.3337+h)
    P23 = oTG.CreatePoint2d(2.4609 , 2.1097+h)
    P24 = oTG.CreatePoint2d(1.2730, 2.4100+h)

    P25 = oTG.CreatePoint2d(4.4022, 1.5725+h)
    P26 = oTG.CreatePoint2d(3.0766, 1.7772+h)
    P27 = oTG.CreatePoint2d(5.5590 , 2.2514+h)

    P28 = oTG.CreatePoint2d(10.0282, 6.5381+h)
    P29 = oTG.CreatePoint2d(8.5294, 5.7909+h)
    P30 = oTG.CreatePoint2d(11.5968, 5.9516+h)

    P31 = oTG.CreatePoint2d(15.3081, 2.5192+h)
    P32 = oTG.CreatePoint2d(14.8280, 2.8310+h)
    P33 = oTG.CreatePoint2d(15.8700, 2.4100+h)

    P34 = oTG.CreatePoint2d(18.4121 , 2.3221+h)
    P35 = oTG.CreatePoint2d(18.5 , 2.11+h)
    P36 = oTG.CreatePoint2d(18.2, 2.41+h)

    oArc = oSketch.SketchArcs
    oLines = oSketch.SketchLines
    
    oArc1 = oArc.AddByThreePoints(P1, P2, P3)#
    oLines1 = oLines.AddByTwoPoints(oArc1.EndSketchPoint, P6)
    oArc2 = oArc.AddByThreePoints(oLines1.EndSketchPoint, P4, P5)
    oLines2 = oLines.AddByTwoPoints(oArc2.StartSketchPoint, P8)
    oArc3 = oArc.AddByThreePoints(oLines2.EndSketchPoint, P7, P9)
    oLines3 = oLines.AddByTwoPoints(oArc3.EndSketchPoint, P11)
    oArc4 = oArc.AddByThreePoints(oLines3.EndSketchPoint, P10, P12)
    oLines4 = oLines.AddByTwoPoints(oArc4.StartSketchPoint, P14)
    oArc5 = oArc.AddByThreePoints(oLines4.EndSketchPoint, P13, P15)
    oLines5 = oLines.AddByTwoPoints(oArc5.EndSketchPoint, P17)
    oArc6 = oArc.AddByThreePoints(oLines5.EndSketchPoint, P16, P18)
    oLines6 = oLines.AddByTwoPoints(oArc1.StartSketchPoint, P21)
    oArc7 = oArc.AddByThreePoints(oLines6.EndSketchPoint, P19, P20)
    oLines7 = oLines.AddByTwoPoints(oArc7.StartSketchPoint, P24)
    oArc8 = oArc.AddByThreePoints(oLines7.EndSketchPoint, P22, P23)
    oLines8 = oLines.AddByTwoPoints(oArc8.StartSketchPoint, P26)
    oArc9 = oArc.AddByThreePoints(oLines8.EndSketchPoint, P25, P27)
    oLines9 = oLines.AddByTwoPoints(oArc9.EndSketchPoint, P29)
    oArc10 = oArc.AddByThreePoints(oLines9.EndSketchPoint, P28, P30)
    oLines10 = oLines.AddByTwoPoints(oArc10.StartSketchPoint, P32)
    oArc11 = oArc.AddByThreePoints(oLines10.EndSketchPoint, P31, P33)
    oLines11 = oLines.AddByTwoPoints(oArc11.EndSketchPoint, P36)
    oArc12 = oArc.AddByThreePoints(oLines11.EndSketchPoint, P34, P35)
    oLines12 = oLines.AddByTwoPoints(oArc12.StartSketchPoint, oArc6.EndSketchPoint)
    oProfile = oSketch.Profiles.AddForSolid()
    oExtFeature = oPartDoc.ComponentDefinition.Features.ExtrudeFeatures.AddByDistanceExtent(oProfile, 0.1, 20995, 20481)
    oApp.ActiveView.GoHome()


#height_input = float(input("Enter height: "))
profile_offset = 2.41 #will be taken as an input
actual_offset = 2.41

combiner(profile_offset - actual_offset)