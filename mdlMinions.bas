Attribute VB_Name = "mdlMinions"
Option Explicit

Sub subMinions()
Attribute subMinions.VB_ProcData.VB_Invoke_Func = "m\n14"

    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim strGrp() As String, i As Integer, j As Integer
    i = -1
    Select Case Int(Rnd * 10)
        Case Is >= 5
            j = 0
        Case Else
            j = 10
    End Select
    
    Dim obj As Object
    'Body
    ws.Shapes.AddShape(msoShapeFlowchartDelay, 70, 30, 40, 40).Select
    Set obj = Selection.ShapeRange
    obj.IncrementRotation -90
    obj.Fill.ForeColor.RGB = RGB(255, 217, 102)
    obj.Line.Visible = False
    i = i + 1
    ReDim Preserve strGrp(i)
    strGrp(i) = obj.Name
    
    ws.Shapes.AddShape(msoShapeRectangle, 70, 70, 40, 20 + j).Select
    Set obj = Selection.ShapeRange
    obj.Fill.ForeColor.RGB = RGB(255, 217, 102)
    obj.Line.Visible = False
    i = i + 1
    ReDim Preserve strGrp(i)
    strGrp(i) = obj.Name

    'Left Hand
    Select Case Int(Rnd * 10)
        Case Is >= 5
            ws.Shapes.AddShape(msoShapeBlockArc, 58, 78 + j, 26, 22).Select
            Set obj = Selection.ShapeRange
            obj.Fill.ForeColor.RGB = RGB(255, 217, 102)
            obj.IncrementRotation -90
            obj.Adjustments.Item(3) = 0.3
            obj.Line.Visible = False
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name

        Case Is < 5
            ws.Shapes.AddShape(msoShapeHeart, 50, 52 + j, 8, 30).Select
            Set obj = Selection.ShapeRange
            obj.Fill.ForeColor.RGB = RGB(64, 64, 64)
            obj.IncrementRotation -45
            obj.Line.Visible = False
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name
        
            ws.Shapes.AddShape(msoShapeRectangle, 50, 74 + j, 30, 8).Select
            Set obj = Selection.ShapeRange
            obj.Fill.ForeColor.RGB = RGB(255, 217, 102)
            obj.IncrementRotation 45
            obj.Line.Visible = False
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name
        
            ws.Shapes.AddShape(msoShapeOval, 48, 61 + j, 12, 12).Select
            Set obj = Selection.ShapeRange
            obj.Fill.ForeColor.RGB = RGB(64, 64, 64)
            obj.Line.Visible = False
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name
            
    End Select

    'Right Hand
    Select Case Int(Rnd * 10)
        Case Is >= 5
            ws.Shapes.AddShape(msoShapeBlockArc, 96, 78 + j, 26, 22).Select
            Set obj = Selection.ShapeRange
            obj.Fill.ForeColor.RGB = RGB(255, 217, 102)
            obj.IncrementRotation 90
            obj.Adjustments.Item(3) = 0.3
            obj.Line.Visible = False
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name
        
        Case Is < 5
            ws.Shapes.AddShape(msoShapeHeart, 122, 52 + j, 8, 30).Select
            Set obj = Selection.ShapeRange
            obj.Fill.ForeColor.RGB = RGB(64, 64, 64)
            obj.IncrementRotation 45
            obj.Line.Visible = False
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name
        
            ws.Shapes.AddShape(msoShapeRectangle, 100, 74 + j, 30, 8).Select
            Set obj = Selection.ShapeRange
            obj.Fill.ForeColor.RGB = RGB(255, 217, 102)
            obj.IncrementRotation -45
            obj.Line.Visible = False
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name
        
            ws.Shapes.AddShape(msoShapeOval, 120, 61 + j, 12, 12).Select
            Set obj = Selection.ShapeRange
            obj.Fill.ForeColor.RGB = RGB(64, 64, 64)
            obj.Line.Visible = False
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name
        
    End Select

    'Cloth
    ws.Shapes.AddShape(msoShapeRectangle, 75, 80 + j, 30, 10).Select
    Set obj = Selection.ShapeRange
    obj.Fill.ForeColor.RGB = RGB(68, 114, 196)
    obj.Line.Visible = False
    i = i + 1
    ReDim Preserve strGrp(i)
    strGrp(i) = obj.Name

    ws.Shapes.AddShape(msoShapeRectangle, 70, 76 + j, 10, 5).Select
    Set obj = Selection.ShapeRange
    obj.Fill.ForeColor.RGB = RGB(68, 114, 196)
    obj.IncrementRotation 45
    obj.Line.Visible = False
    i = i + 1
    ReDim Preserve strGrp(i)
    strGrp(i) = obj.Name

    ws.Shapes.AddShape(msoShapeRectangle, 100, 76 + j, 10, 5).Select
    Set obj = Selection.ShapeRange
    obj.Fill.ForeColor.RGB = RGB(68, 114, 196)
    obj.IncrementRotation -45
    obj.Line.Visible = False
    i = i + 1
    ReDim Preserve strGrp(i)
    strGrp(i) = obj.Name

    ws.Shapes.AddShape(msoShapeFlowchartDelay, 80, 80 + j, 20, 40).Select
    Set obj = Selection.ShapeRange
    obj.IncrementRotation 90
    obj.Fill.ForeColor.RGB = RGB(68, 114, 196)
    obj.Line.Visible = False
    i = i + 1
    ReDim Preserve strGrp(i)
    strGrp(i) = obj.Name

    ws.Shapes.AddShape(msoShapeFlowchartDelay, 85, 84 + j, 10, 16).Select
    Set obj = Selection.ShapeRange
    obj.IncrementRotation 90
    obj.Fill.ForeColor.RGB = RGB(68, 114, 196)
    obj.Line.ForeColor.RGB = RGB(255, 255, 255)
    i = i + 1
    ReDim Preserve strGrp(i)
    strGrp(i) = obj.Name
    
    'Leg
    ws.Shapes.AddShape(msoShapeRectangle, 80, 105 + j, 6, 10).Select
    Set obj = Selection.ShapeRange
    obj.Fill.ForeColor.RGB = RGB(68, 114, 196)
    obj.Line.Visible = False
    i = i + 1
    ReDim Preserve strGrp(i)
    strGrp(i) = obj.Name
    
    ws.Shapes.AddShape(msoShapeFlowchartManualInput, 72, 114 + j, 14, 8).Select
    Set obj = Selection.ShapeRange
    obj.Fill.ForeColor.RGB = RGB(64, 64, 64)
    obj.Line.Visible = False
    i = i + 1
    ReDim Preserve strGrp(i)
    strGrp(i) = obj.Name

    ws.Shapes.AddShape(msoShapeRectangle, 94, 105 + j, 6, 10).Select
    Set obj = Selection.ShapeRange
    obj.Fill.ForeColor.RGB = RGB(68, 114, 196)
    obj.Line.Visible = False
    i = i + 1
    ReDim Preserve strGrp(i)
    strGrp(i) = obj.Name

    ws.Shapes.AddShape(msoShapeFlowchartManualInput, 94, 114 + j, 14, 8).Select
    Set obj = Selection.ShapeRange
    obj.Fill.ForeColor.RGB = RGB(64, 64, 64)
    obj.Line.Visible = False
    obj.Flip msoFlipHorizontal
    i = i + 1
    ReDim Preserve strGrp(i)
    strGrp(i) = obj.Name

    'Eye
    ws.Shapes.AddShape(msoShapeRectangle, 70, 50, 40, 8).Select
    Set obj = Selection.ShapeRange
    obj.Fill.ForeColor.RGB = RGB(64, 64, 64)
    obj.Line.Visible = False
    i = i + 1
    ReDim Preserve strGrp(i)
    strGrp(i) = obj.Name
    
    Select Case Int(Rnd * 10)
        Case Is >= 5
            ws.Shapes.AddShape(msoShapeOval, 77, 40, 26, 26).Select
            Set obj = Selection.ShapeRange
            obj.Fill.ForeColor.RGB = RGB(255, 255, 255)
            obj.Line.Visible = False
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name

            ws.Shapes.AddShape(msoShapeOval, 85, 49, 10, 10).Select
            Set obj = Selection.ShapeRange
            obj.Fill.ForeColor.RGB = RGB(90, 40, 10)
            obj.Line.Visible = False
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name

            'Eyelid
            Select Case Int(Rnd * 10)
                Case Is >= 5
                    ws.Shapes.AddShape(msoShapePie, 78, 41, 24, 22).Select
                    Set obj = Selection.ShapeRange
                    obj.Adjustments.Item(1) = 180
                    obj.Adjustments.Item(2) = 0
                    obj.Fill.ForeColor.RGB = RGB(255, 217, 102)
                    obj.Line.Visible = False
                    i = i + 1
                    ReDim Preserve strGrp(i)
                    strGrp(i) = obj.Name
            End Select

            ws.Shapes.AddShape(msoShapeOval, 77, 40, 26, 26).Select
            Set obj = Selection.ShapeRange
            obj.Fill.Visible = False
            obj.Line.Weight = 3
            obj.Line.ForeColor.RGB = RGB(127, 127, 127)
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name

        Case Is < 5
            ws.Shapes.AddShape(msoShapeOval, 73, 45, 17, 17).Select
            Set obj = Selection.ShapeRange
            obj.Fill.ForeColor.RGB = RGB(255, 255, 255)
            obj.Line.Visible = False
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name

            ws.Shapes.AddShape(msoShapeOval, 90, 45, 17, 17).Select
            Set obj = Selection.ShapeRange
            obj.Fill.ForeColor.RGB = RGB(255, 255, 255)
            obj.Line.Visible = False
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name

            ws.Shapes.AddShape(msoShapeOval, 79, 50, 6, 6).Select
            Set obj = Selection.ShapeRange
            obj.Fill.ForeColor.RGB = RGB(90, 40, 10)
            obj.Line.Visible = False
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name

            ws.Shapes.AddShape(msoShapeOval, 95, 50, 6, 6).Select
            Set obj = Selection.ShapeRange
            obj.Fill.ForeColor.RGB = RGB(90, 40, 10)
            obj.Line.Visible = False
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name
            
            'Eyelid
            Select Case Int(Rnd * 10)
                Case Is >= 5
                    ws.Shapes.AddShape(msoShapePie, 74, 43, 14, 18).Select
                    Set obj = Selection.ShapeRange
                    obj.Adjustments.Item(1) = 180
                    obj.Adjustments.Item(2) = 0
                    obj.Fill.ForeColor.RGB = RGB(255, 217, 102)
                    obj.Line.Visible = False
                    i = i + 1
                    ReDim Preserve strGrp(i)
                    strGrp(i) = obj.Name

                    ws.Shapes.AddShape(msoShapePie, 92, 43, 14, 18).Select
                    Set obj = Selection.ShapeRange
                    obj.Adjustments.Item(1) = 180
                    obj.Adjustments.Item(2) = 0
                    obj.Fill.ForeColor.RGB = RGB(255, 217, 102)
                    obj.Line.Visible = False
                    i = i + 1
                    ReDim Preserve strGrp(i)
                    strGrp(i) = obj.Name
            End Select

            ws.Shapes.AddShape(msoShapeOval, 73, 45, 17, 17).Select
            Set obj = Selection.ShapeRange
            obj.Fill.Visible = False
            obj.Line.Weight = 3
            obj.Line.ForeColor.RGB = RGB(127, 127, 127)
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name

            ws.Shapes.AddShape(msoShapeOval, 90, 45, 17, 17).Select
            Set obj = Selection.ShapeRange
            obj.Fill.Visible = False
            obj.Line.Weight = 3
            obj.Line.ForeColor.RGB = RGB(127, 127, 127)
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name

    End Select

    'Mouth
    Select Case Int(Rnd * 10)
        Case Is >= 8
            ws.Shapes.AddShape(msoShapeArc, 90, 64, 10, 6).Select
            Set obj = Selection.ShapeRange
            obj.Adjustments.Item(1) = 180
            obj.IncrementRotation 180
            obj.Line.ForeColor.RGB = RGB(64, 64, 64)
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name
       
       Case Is >= 5
            ws.Shapes.AddShape(msoShapeArc, 90, 69, 10, 6).Select
            Set obj = Selection.ShapeRange
            obj.Adjustments.Item(1) = 180
            obj.Line.ForeColor.RGB = RGB(64, 64, 64)
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name
        
        Case Is >= 3
            ws.Shapes.AddShape(msoShapeTrapezoid, 80, 70, 20, 8).Select
            Set obj = Selection.ShapeRange
            obj.Fill.ForeColor.RGB = RGB(255, 0, 102)
            obj.Line.Visible = False
            obj.IncrementRotation 180
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name
    
        Case Else
            ws.Shapes.AddShape(msoShapeTrapezoid, 85, 70, 10, 8).Select
            Set obj = Selection.ShapeRange
            obj.Fill.ForeColor.RGB = RGB(255, 0, 102)
            obj.Line.Visible = False
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name
    End Select
    
    'Hair
    Select Case Int(Rnd * 10)
        Case Is >= 8
            ws.Shapes.AddShape(msoShapeArc, 83, 30, 7, 3).Select
            Set obj = Selection.ShapeRange
            obj.Adjustments.Item(1) = 180
            obj.IncrementRotation -22
            obj.Line.ForeColor.RGB = RGB(64, 64, 64)
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name

            ws.Shapes.AddShape(msoShapeArc, 97, 30, 7, 3).Select
            Set obj = Selection.ShapeRange
            obj.Adjustments.Item(1) = 180
            obj.IncrementRotation 22
            obj.Line.ForeColor.RGB = RGB(64, 64, 64)
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name

        Case Is >= 5
            ws.Shapes.AddShape(msoShapeArc, 86, 26, 5, 3).Select
            Set obj = Selection.ShapeRange
            obj.Adjustments.Item(1) = 180
            obj.IncrementRotation 40
            obj.Line.ForeColor.RGB = RGB(64, 64, 64)
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name

            ws.Shapes.AddShape(msoShapeArc, 95, 26, 5, 3).Select
            Set obj = Selection.ShapeRange
            obj.Adjustments.Item(1) = 180
            obj.IncrementRotation -40
            obj.Line.ForeColor.RGB = RGB(64, 64, 64)
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name
        
        Case Else
            ws.Shapes.AddShape(msoShapeArc, 85, 26, 3, 2).Select
            Set obj = Selection.ShapeRange
            obj.Adjustments.Item(1) = 180
            obj.IncrementRotation 90
            obj.Line.ForeColor.RGB = RGB(64, 64, 64)
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name
            
            ws.Shapes.AddShape(msoShapeArc, 90, 26, 3, 2).Select
            Set obj = Selection.ShapeRange
            obj.Adjustments.Item(1) = 180
            obj.IncrementRotation 90
            obj.Line.ForeColor.RGB = RGB(64, 64, 64)
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name

            ws.Shapes.AddShape(msoShapeArc, 95, 26, 3, 2).Select
            Set obj = Selection.ShapeRange
            obj.Adjustments.Item(1) = 180
            obj.IncrementRotation 90
            obj.Line.ForeColor.RGB = RGB(64, 64, 64)
            i = i + 1
            ReDim Preserve strGrp(i)
            strGrp(i) = obj.Name


    End Select
    
    'Grouping & Select
    ws.Shapes.Range(strGrp).Group.Select
    Selection.ShapeRange.Top = Int((300 - 10 + 1) * Rnd + 10)
    Selection.ShapeRange.Left = Int((800 - 10 + 1) * Rnd + 10)
    Selection.ShapeRange.Rotation = Int((360 - 0 + 1) * Rnd + 0)
    
    Application.ScreenUpdating = True
    
End Sub

