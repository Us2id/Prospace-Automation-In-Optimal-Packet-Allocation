import win32com.client as wc

def node_sketch(x, y, oTG, pw):
    p = (-1)**pw
    centre = oTG.CreatePoint2d(p*x, -y)
    p1 = oTG.CreatePoint2d(p*(x-.36), -(y-.36))
    p2 = oTG.CreatePoint2d(p*(x-.16), -(y-.16))
    p3 = oTG.CreatePoint2d(p*(x+.16), -(y-.16))
    p4 = oTG.CreatePoint2d(p*(x+.36), -(y-.36))
    return centre, p1, p2, p3, p4
def draw_sketch(centre, p1, p2, p3, p4, oSketch):
    oSketch.SketchLines.AddAsTwoPointCenteredRectangle(centre, p2)
    oLines1 = oSketch.SketchLines.AddByTwoPoints(p1, p2)
    oLines2 = oSketch.SketchLines.AddByTwoPoints(oLines1.EndSketchPoint, p3)
    oLines3 = oSketch.SketchLines.AddByTwoPoints(oLines2.EndSketchPoint, p4)
    oLines4 = oSketch.SketchLines.AddByTwoPoints(oLines3.EndSketchPoint, oLines1.StartSketchPoint)
def rotate(Sketch, oApp, angle):
    objCollection = oApp.TransientObjects.CreateObjectCollection()
    angle = angle * (3.141592653589793 / 180)
    for line in Sketch.SketchLines:
        objCollection.Add(line)
    for line in Sketch.SketchArcs:
        objCollection.Add(line)
    Sketch.RotateSketchObjects(objCollection, oApp.TransientGeometry.CreatePoint2d(0, 0), angle)

def combiner(l):
    oApp = wc.GetActiveObject('Inventor.Application')
    #oPartDoc = oApp.Documents.Add(12290, oApp.FileManager.GetTemplateFile(12290, 8962, 9729, "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"))
    oPartDoc = oApp.Documents.Add(12290, oApp.FileManager.GetTemplateFile(12290, 8962))
    oSketch = oPartDoc.ComponentDefinition.Sketches.Add(oPartDoc.ComponentDefinition.WorkPlanes.Item(3))
    oTG = oApp.TransientGeometry

#draing the rod with curves
    #lower profile
    P1 = oTG.CreatePoint2d(0, 0)
    P2 = oTG.CreatePoint2d(-0.07, 0)
    P3 = oTG.CreatePoint2d(-0.07, 1.65)
    P4 = oTG.CreatePoint2d(0.0809, 2.4259)
    P5 = oTG.CreatePoint2d(.5116, 3.0886)
    P6 = oTG.CreatePoint2d(3.5480, 6.23)
    P7 = oTG.CreatePoint2d(4.1073, 7.7240)
    P8 = oTG.CreatePoint2d(3.3959, 9.1518)
    P9 = oTG.CreatePoint2d(-.1809, 12.1546)
    P10 = oTG.CreatePoint2d(-.8824 , 13.3473)
    P11 = oTG.CreatePoint2d(-.6757, 14.7155)
    P12 = oTG.CreatePoint2d(-.3060, 15.4075)
    P13 = oTG.CreatePoint2d(-.1299, 15.8642)
    P14 = oTG.CreatePoint2d(-0.07, 16.35)
    P15 = oTG.CreatePoint2d(-0.07, 17)
    P16 = oTG.CreatePoint2d(0, 17)
    #upper profile
    P17 = oTG.CreatePoint2d(0, 16.35)
    P18 = oTG.CreatePoint2d(-0.0620, 15.8472)   
    P19 = oTG.CreatePoint2d(-.2443, 15.3745)
    P20 = oTG.CreatePoint2d(-.6140, 14.6825)
    P21 = oTG.CreatePoint2d(-.8137, 13.3606)
    P22 = oTG.CreatePoint2d(-.1359, 12.2082)
    P23 = oTG.CreatePoint2d(3.4409 , 9.2054)
    P24 = oTG.CreatePoint2d(4.1772, 7.7276)
    P25 = oTG.CreatePoint2d(3.5984, 6.1814)
    P26 = oTG.CreatePoint2d(.5620, 3.04)
    P27 = oTG.CreatePoint2d(.1458 , 2.3996)
    P28 = oTG.CreatePoint2d(0, 1.65)
    
    oArc = oSketch.SketchArcs
    oLines = oSketch.SketchLines

    oLines1 = oLines.AddByTwoPoints(P1, P2)
    oLines2 = oLines.AddByTwoPoints(oLines1.EndSketchPoint, P3)
    oArc1 = oArc.AddByThreePoints(oLines2.EndSketchPoint, P4, P5)
    oLines3 = oLines.AddByTwoPoints(oArc1.StartSketchPoint, P6)
    oArc2 = oArc.AddByThreePoints(oLines3.EndSketchPoint, P7, P8)
    oLines4 = oLines.AddByTwoPoints(oArc2.EndSketchPoint, P9)
    oArc3 = oArc.AddByThreePoints(oLines4.EndSketchPoint, P10, P11)
    oLines5 = oLines.AddByTwoPoints(oArc3.StartSketchPoint, P12)
    oArc4 = oArc.AddByThreePoints(oLines5.EndSketchPoint, P13, P14)
    oLines6 = oLines.AddByTwoPoints(oArc4.EndSketchPoint, P15)
    oLines7 = oLines.AddByTwoPoints(oLines6.EndSketchPoint, P16)
    oLines8 = oLines.AddByTwoPoints(oLines7.EndSketchPoint, P17)
    oArc5 = oArc.AddByThreePoints(oLines8.EndSketchPoint, P18, P19)
    oLines9 = oLines.AddByTwoPoints(oArc5.StartSketchPoint, P20)
    oArc6 = oArc.AddByThreePoints(oLines9.EndSketchPoint, P21, P22)
    oLines10 = oLines.AddByTwoPoints(oArc6.EndSketchPoint, P23)
    oArc7 = oArc.AddByThreePoints(oLines10.EndSketchPoint, P24, P25)
    oLines11 = oLines.AddByTwoPoints(oArc7.StartSketchPoint, P26)
    oArc8 = oArc.AddByThreePoints(oLines11.EndSketchPoint, P27, P28)
    oLines12 = oLines.AddByTwoPoints(oArc8.EndSketchPoint, oLines1.StartSketchPoint)
    rotate(oSketch, oApp, 90)
    oProfile = oSketch.Profiles.AddForSolid()
    oExtFeature = oPartDoc.ComponentDefinition.Features.ExtrudeFeatures.AddByDistanceExtent(oProfile, l, 20995, 20481)

    oApp.ActiveView.GoHome()
    
    num_notches = int((l-1)/50) + 1
    division_len = (l-1)/num_notches

#making notches in the rod
    oFace1 = oPartDoc.ComponentDefinition.SurfaceBodies.Item(1).Faces.Item(12)
    oEdge1 = oFace1.Edges.Item(1)
    oVertex1 = oEdge1.StartVertex
    oSketch1 = oPartDoc.ComponentDefinition.Sketches.AddWithOrientation(oFace1, oEdge1, True, True, oVertex1)
    for i in range(num_notches):
        x, y = .5 + i*division_len, .36
        centre, p1, p2, p3, p4 = node_sketch(x, y, oTG, 2)
        draw_sketch(centre, p1, p2, p3, p4, oSketch1)
    x, y = l-.5, .36
    centre, p1, p2, p3, p4 = node_sketch(x, y, oTG, 2)
    draw_sketch(centre, p1, p2, p3, p4, oSketch1)

    oFace2 = oPartDoc.ComponentDefinition.SurfaceBodies.Item(1).Faces.Item(20)
    oEdge2 = oFace2.Edges.Item(2)
    oVertex2 = oEdge2.StartVertex
    oSketch2 = oPartDoc.ComponentDefinition.Sketches.AddWithOrientation(oFace2, oEdge2, True, False, oVertex2)
    for i in range(num_notches):
        x, y = .5 + i*division_len, .36
        centre, p1, p2, p3, p4 = node_sketch(x, y, oTG, 3)
        draw_sketch(centre, p1, p2, p3, p4, oSketch2)
    x, y = l-.5, .36
    centre, p1, p2, p3, p4 = node_sketch(x, y, oTG, 3)
    draw_sketch(centre, p1, p2, p3, p4, oSketch2)

    solid_profile1 = oSketch2.Profiles.AddforSolid()
    ext_solid_def1 = oPartDoc.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(solid_profile1, 20482)
    ext_solid_def1.SetDistanceExtent(0.07, 20994)

    solid_profile1 = oSketch1.Profiles.AddforSolid()
    ext_solid_def2 = oPartDoc.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(solid_profile1, 20482)
    ext_solid_def2.SetDistanceExtent(0.07, 20994)

    oPartDoc.ComponentDefinition.Features.ExtrudeFeatures.Add(ext_solid_def1)
    oPartDoc.ComponentDefinition.Features.ExtrudeFeatures.Add(ext_solid_def2)


#height_input = float(input("Enter height: "))
length = 147
combiner(length)