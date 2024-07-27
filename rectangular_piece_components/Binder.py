import win32com.client as wc

def rotate(Sketch, oApp, angle):
    objCollection = oApp.TransientObjects.CreateObjectCollection()
    angle = angle * (3.141592653589793 / 180)
    for line in Sketch.SketchLines:
        objCollection.Add(line)
    for line in Sketch.SketchArcs:
        objCollection.Add(line)
    Sketch.RotateSketchObjects(objCollection, oApp.TransientGeometry.CreatePoint2d(0, 0), angle)
def combiner(h):
    oApp = wc.GetActiveObject('Inventor.Application')

    oPartDoc = oApp.Documents.Add(12290, oApp.FileManager.GetTemplateFile(12290, 8962))

    oSketch = oPartDoc.ComponentDefinition.Sketches.Add(oPartDoc.ComponentDefinition.WorkPlanes.Item(2))

    oTG = oApp.TransientGeometry

    oSkPnts = oSketch.SketchPoints
    #inner circle
    oSkPnts.Add(oTG.CreatePoint2d(0, 0), False)     #1
    oSkPnts.Add(oTG.CreatePoint2d(0, -0.05), False) #2
    oSkPnts.Add(oTG.CreatePoint2d(0, 0.05), False)  #3
    #outer circle
    oSkPnts.Add(oTG.CreatePoint2d(0, -0.15), False) #4
    oSkPnts.Add(oTG.CreatePoint2d(0, 0.15), False)  #5
    #inner line
    oSkPnts.Add(oTG.CreatePoint2d(-1.35, -0.05), False) #6
    oSkPnts.Add(oTG.CreatePoint2d(-1.35, 0.05), False)  #7
    #outer line
    oSkPnts.Add(oTG.CreatePoint2d(-1.35, -0.15), False) #8
    oSkPnts.Add(oTG.CreatePoint2d(-1.35, 0.15), False)  #9

    oArc = oSketch.SketchArcs
    oArc1 = oArc.AddByCenterStartEndPoint(oSkPnts(1), oSkPnts(2), oSkPnts(3))
    oArc2 = oArc.AddByCenterStartEndPoint(oSkPnts(1), oSkPnts(4), oSkPnts(5))

    oLines = oSketch.SketchLines
    oLine1 = oLines.AddByTwoPoints(oSkPnts(6), oSkPnts(2))
    oLine2 = oLines.AddByTwoPoints(oSkPnts(7), oSkPnts(3))
    oLine3 = oLines.AddByTwoPoints(oSkPnts(8), oSkPnts(4))
    oLine4 = oLines.AddByTwoPoints(oSkPnts(9), oSkPnts(5))
    oLine5 = oLines.AddByTwoPoints(oSkPnts(6), oSkPnts(8))
    oLine6 = oLines.AddByTwoPoints(oSkPnts(7), oSkPnts(9))

    #rotate(oSketch, oApp, 180)

    oProfile = oSketch.Profiles.AddForSolid()
    oExtFeature = oPartDoc.ComponentDefinition.Features.ExtrudeFeatures.AddByDistanceExtent(oProfile, h, 20995, 20481)

    edge1 = oPartDoc.ComponentDefinition.SurfaceBodies.Item(1).Faces.Item(3).Edges.Item(2)
    edge2 = oPartDoc.ComponentDefinition.SurfaceBodies.Item(1).Faces.Item(3).Edges.Item(4)
    edge3 = oPartDoc.ComponentDefinition.SurfaceBodies.Item(1).Faces.Item(7).Edges.Item(2)
    edge4 = oPartDoc.ComponentDefinition.SurfaceBodies.Item(1).Faces.Item(7).Edges.Item(4)

    edge_collection = oApp.TransientObjects.CreateEdgeCollection()
    fillet_def = oPartDoc.ComponentDefinition.Features.FilletFeatures.CreateFilletDefinition()

    edge_collection.Add(edge1)
    edge_collection.Add(edge2)
    edge_collection.Add(edge3)
    edge_collection.Add(edge4)

    fillet_def.AddConstantRadiusEdgeSet(edge_collection, 0.1)
    oPartDoc.ComponentDefinition.Features.FilletFeatures.Add(fillet_def)

    oApp.ActiveView.GoHome()

#height_input = float(input("Enter height: "))

combiner(37.18)