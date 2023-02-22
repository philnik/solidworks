
Part = sw.ActiveDoc

Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, ARG_NULL, 0)

Part.SketchManager.InsertSketch(True)
sketch=Part.SketchManager.ActiveSketch
Part.ClearSelection2
skSegment = Part.SketchManager.CreateCircle(0, 0, 0, 0.0, 0.125, 0)
#Part.ClearSelection2
Part.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", -0.125750978545243, 0, 6.76383723033451E-02, False, 0, ARG_NULL, 0)
#sw.SetUserPreferenceToggle(2,False)
myDisplayDim = Part.AddDimension2(-0.241948077546718, 0, 6.48042967179433E-02)

Part.ClearSelection2
#myDimension = Part.Parameter("D1@Sketch1")
#myDimension.SystemValue = 0.1

Part.ClearSelection2
Part.SketchManager.InsertSketch(True)
boolstatus = Part.Extension.SelectByID2(sketch.Name, "SKETCH", 0, 0, 0, False, 0, ARG_NULL, 0)

length=0.5
myFeature = Part.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, length, 0.01, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, True, True, True, 0, 0, False)


