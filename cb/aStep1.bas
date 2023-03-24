Attribute VB_Name = "aStep1"
Option Private Module

'Step1
Sub GetSTL(strSTLFilePath As String)
'1. merge STL file in 'STL' Layer
    'Call Document.MergeFile("C:\Users\user\Desktop\기본설정\TEST\Osstem TS,GS Standard ver.1.esp")
    'Document.Refresh
    Dim lyrObjectInitial As Esprit.Layer
    Dim lyrObject As Esprit.Layer
    Set lyrObjectInitial = Document.ActiveLayer
    
    For Each lyrObject In Esprit.Document.Layers
        With lyrObject
        If (.Name = "STL") Then
            Document.ActiveLayer = lyrObject
        End If
        End With
    Next
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = "STL" And (.GraphicObjectType = espSTL_Model)) Then
            If (.Key > 0) Then
                Document.GraphicsCollection.Remove (graphicObject.GraphicsCollectionIndex)
                Call MsgBox("Previous STL model has been deleted.", , "CAM Automation")
                Document.Refresh
            End If
        End If
        End With
    Next
    
    Call Document.MergeFile(strSTLFilePath)
    'Call Document.MergeFile("Select STL File")
    
    
    Document.Refresh

End Sub

Public Function CheckSTLInTheCircle() As Boolean
'Is The STL in the Circle?

Call GetPartProfileSTL

'Function IntersectCircleAndArcsSegments(ByRef cBound As Esprit.Circle, _
'            ByRef go2 As Esprit.graphicObject) As Esprit.Point

    Dim cBound As Esprit.Circle
    Dim go2 As Esprit.graphicObject
    Dim pntIntersect As Esprit.Point
    Dim layerObject As Esprit.Layer
    Dim graphicObject As Esprit.graphicObject
        
    Dim crl As Esprit.Circle
    For Each crl In Esprit.Document.Circles
        With crl
        If (.Layer.Name = "기본값") Then
            If (.Key > 0) Then
                Set cBound = crl
                .Grouped = True
                .Layer.Visible = True
                Set layerObject = .Layer
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next

    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = "STL" And (.GraphicObjectType = espSegment Or .GraphicObjectType = espArc)) Then
            If (.Key > 0) Then
                Set go2 = graphicObject
                .Grouped = True
                .Layer.Visible = True
                Set pntIntersect = IntersectCircleAndArcsSegments(cBound, go2)
                Call Document.GraphicsCollection.Remove(go2.GraphicsCollectionIndex)
                Document.Refresh
                
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
    
    'Dim bInTheCircle As Boolean
    CheckSTLInTheCircle = (pntIntersect Is Nothing)
    
    If CheckSTLInTheCircle Then
        Call Document.GraphicsCollection.Remove(cBound.GraphicsCollectionIndex)
    End If
    
    Document.Refresh

'IntersectCircleAndArcsSegments(

End Function

Sub SelectSTL_Model()
'4. move the selected STL feature
    Dim mGraphicObject As Esprit.graphicObject
    Dim i As Long
    
     'Check the existing selection
    'Document.OpenUndoTransaction
    Dim stl As Esprit.STL_Model
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = "STL" And (.GraphicObjectType = espSTL_Model)) Then
            If (.Key > 0) Then
                Set stl = graphicObject
                Call MoveSTL_Step1(stl)
                Document.Refresh
            End If
        End If
        End With
    Next
    
    'Call Document.CloseUndoTransaction(True)
    Set mGraphicObject = Nothing

    Dim nCount As Integer
    'nCount = Document.Group.Count
    'Call Document.SelectionSets.Item(1).Rotate(l, -90, 1)
    nCount = Document.Group.Count
    
    Call Application.OutputWindow.Text("PartStockLength(Before): " & CStr(Document.LatheMachineSetup.PartStockLength) & vbCrLf)
    Document.LatheMachineSetup.PartStockLength = Round(GetCutOffXRightEnd(), 2)
    Call Application.OutputWindow.Text("PartStockLength(Updated): " & CStr(Document.LatheMachineSetup.PartStockLength) & vbCrLf)
    
'    Dim sitm As Esprit.SelectionSet
''    sitm.Add (Document.Group.Item(1))
'    Set sitm = Document.SelectionSets.Add("STL")
End Sub


'This function will invert the input ptop orientation
Sub MoveSTL_Step1(ByRef stl As Esprit.STL_Model)
    
    If stl Is Nothing Then Exit Sub 'check if the object is valid
    
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("Temp")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("Temp")
    End With

    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(stl)  'Add the stl_model to the selection object
    End With
    
    ' Call mSelection.Translate(5, 0, 0)
    '1. Rotate
    Dim radian As Double
    Dim degree As Double
    Dim radians_angle As Double
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    '3)  Y축기준 -90도 회전
    'radian = -90 * PI / 180
    Call TurnSTL("Y,-90")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    '4)  X축기준 (Hexa: -30도 회전 / Square: - 45도 회전)
    Dim strTurning As String
        
    Dim cCondition As Boolean
    cCondition = True
    strTurning = InputBox("Select (H)exa/(O)cta/(S)quare or input (X,-30)", "CAM Automation - To transform the STL", "H")
    Do While cCondition
        Select Case Left(strTurning, 1)
        Case "H" 'Hexa
            Call TurnSTL("X,-30")
            cCondition = False
        Case "O" 'Octa
            Call TurnSTL("X,-45")
            cCondition = False
        Case "S" 'Square
            Call TurnSTL("X,-45")
            cCondition = False
        Case "X", "Y" 'manual input
            If TurnSTL(strTurning) > 0 Then
                cCondition = False
            Else
                strTurning = InputBox("Select (H)exa/(O)cta/(S)quare or input (X,-30)", "CAM Automation - To transform the STL", "X,-30")
            End If
        Case Else
            strTurning = InputBox("Select (H)exa/(O)cta/(S)quare or input (X,-30)", "CAM Automation - To transform the STL", "X,-30")
        End Select
    Loop
        
        
'    radian = -30 * PI / 180
'    'Call Document.Lines.Add(iPoint, 0, 0, 1)
'    Call mSelection.Rotate(iLine, radian, 0)
   
'    iLine.Grouped = True
'    Document.Lines.Remove (Document.Lines.Count)
'    Document.Points.Remove (Document.Points.Count)
    
    
'    For Each circleObject In Esprit.Document.Circles
'        With circleObject
'        If (.Layer.Number = 0) Then
'            .Grouped = True
'        End If
'        End With
'    Next
'
    If Document.Circles.Count > 0 Then
        Call Document.Circles.Remove(Document.Circles.Count)
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Translate move to aside axis X
    'Set mRefGraphicObject() = Document.FeatureRecognition.CreatePartProfileShadow(mSelection, Document.Planes.Item("0DEG"), espFeatureChains)
   ' Dim mRefGraphicObject() As Esprit.graphicObject = Document.FeatureRecognition.CreatePartProfileShadow(mSelection, Document.Planes.Item("0DEG"), espFeatureChains)
    Dim mRefGraphicObject() As Esprit.graphicObject
    Dim returnedFC As Esprit.FeatureChain
    'Dim comCurves() As Object, plottedObjects() As Esprit.graphicObject, faults As EspritComBase.ComFaults
    
    Dim startPoint As Esprit.Point
    Dim midPoint As Esprit.Point
    Dim endPoint As Esprit.Point
    
    Dim graphicObject As Esprit.graphicObject
    Dim dLeftEnd As Double
    dLeftEnd = GetSTLXEnd(0)
    '2. Move X
    Call mSelection.Translate(dLeftEnd * (-1) + 0.1, 0, 0)
    
    Dim r(2) As Double
    
    r(1) = GetCutOffXRightEnd
    r(2) = GetSTLXEnd(1)
        
    Dim nCnt As Integer
    nCnt = 0
    
    'Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("Temp")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("Temp")
    End With
    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
    End With
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        'If (.Layer.Number >= 9 And .Layer.Number <= 11 And (.GraphicObjectType = espSolidModel)) Then
        'If (.Layer.Number >= 9 And .Layer.Number <= 11) Then
        'If ((.Layer.Name = "BACK TURNING" Or .Layer.Name = "CUT-OFF" Or .Layer.Name = "CUF-OFF" Or .Layer.Name = "경계소재-1") And .TypeName <> "Operation") Then
        If ((.Layer.Name = "BACK TURNING" Or .Layer.Name = "CUT-OFF" Or .Layer.Name = "CUF-OFF" Or .Layer.Name = "경계소재-1" Or .Layer.Name = "SPECIAL") And .TypeName <> "Operation") Then
            If (.Key > 0) Then
                'Set solidObject = graphicObject
                
                'Call Step2_ConnectionSet(graphicObject, r(1), r(2))
                'nCnt = nCnt + 1
                Call mSelection.Add(graphicObject)  'Add the stl_model to the selection object
                
                
            
            End If
        End If
        End With
    Next
    
    Call mSelection.Translate(r(2) - r(1), 0, 0, 0)
    Document.Refresh
    
    For Each ly In Document.Layers
        'Requested by Kwangho KJNM at 2018.06.
        'If (ly.Name = "BACK TURNING" Or ly.Name = "CUF-OFF" Or ly.Name = "CUT-OFF" Or ly.Name = "경계소재-1" Or ly.Name = "SPECIAL") Then
        If (ly.Name = "BACK TURNING" Or ly.Name = "CUF-OFF" Or ly.Name = "CUT-OFF" Or ly.Name = "경계소재-1") Then
            ly.Visible = True
        End If
    Next
    
    Document.Refresh
    
End Sub
Sub Step1_2()
    Dim rSTLXEnd As Double
    rSTLXEnd = GetSTLXEnd(1)
    If rSTLXEnd = -999 Then
        Call MsgBox("Cannot get the right end of the STL model in STL Layer. Please check it.")
        Exit Sub
    End If
    
    Dim strTolerance As String
    'strTolerance = InputBox("Enter Tolerance For Turning Profile.", "CAM Automation - For Turning Profile", "0.0001")
    strTolerance = 0.0001
    
    Call GetTurningProfile(CDbl(strTolerance))
    Call GetTurningProfile_EditBoundary(rSTLXEnd)
End Sub

Sub Step2_ConnectionSet(ByRef graphicRef As Esprit.graphicObject, ByVal FromX As Double, ByVal ToX As Double)
    
    If graphicRef Is Nothing Then Exit Sub 'check if the object is valid
    
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("Temp")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("Temp")
    End With

    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(graphicRef)  'Add the stl_model to the selection object
    End With
    
    Call mSelection.Translate(ToX - FromX, 0, 0, 0)
    
    Document.Refresh
End Sub



Function GetCutOffXRightEnd() As Double
'Get connection width from Layer 10. Cut Off
        
    On Error Resume Next

    GetCutOffXRightEnd = 0

    On Error GoTo 0
    
    Dim lyrObjectInitial As Esprit.Layer
    Dim lyrObject As Esprit.Layer
    Set lyrObjectInitial = Document.ActiveLayer
    
    Dim sgmtObject As Esprit.Segment
    Dim sgmtCuttOff(2) As Esprit.Segment
    Dim dCuttOffX(3) As Double
    Dim dReturn As Double
    Dim i As Integer
    
    i = 0
    For Each sgmtObject In Esprit.Document.Segments
        With sgmtObject
        If (.Layer.Name = "CUF-OFF" Or .Layer.Name = "CUT-OFF") Then
            'Document.ActiveLayer = .Layer
            'sgmtObject.Grouped = True
'                sgmtCuttOff(Document.Group.Count) = sgmtObject
            i = i + 1
            dCuttOffX(i) = sgmtObject.XStart
        End If
        End With
    Next
    
    'Get X Right End
    dReturn = 0
    If (dCuttOffX(1) > dCuttOffX(2)) Then
        dReturn = dCuttOffX(1)
    Else
        dReturn = dCuttOffX(2)
    End If
    
    'GetCutOffXRightEnd = dCutOffX(3)
    Document.ActiveLayer = lyrObjectInitial
    
    GetCutOffXRightEnd = dReturn
    Document.Refresh
End Function

Function GetSTLXEnd(nDirectionCode As Integer) As Double
'Get connection width from Layer 10. Cut Off
        
'nDirectionCode = 0 : Left / 1 : Right / else : Error
    On Error Resume Next

    GetSTLXEnd = 0

    On Error GoTo 0
    
    'DirectionCode Check
    If Not (nDirectionCode = 0 Or nDirectionCode = 1) Then
        GoTo EndGetSTLXEnd
    End If
    
    
    Dim graphicObject As Esprit.graphicObject
    Dim stlObject As Esprit.STL_Model
    Dim dReturn As Double
    
    Dim lyrObjectInitial As Esprit.Layer
    Dim lyrObject As Esprit.Layer
    Set lyrObjectInitial = Document.ActiveLayer
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = "STL" And (.GraphicObjectType = espSTL_Model)) Then
            If (.Key > 0) Then
                Set stlObject = graphicObject
            End If
        End If
        End With
    Next
    
    If stlObject Is Nothing Then GetSTLXEnd = -1 'check if the object is valid
    
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("Temp")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("Temp")
    End With

    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(stlObject)  'Add the stl_model to the selection object
    End With
    
    
    Dim mRefGraphicObject() As Esprit.graphicObject
    Dim returnedFC As Esprit.FeatureChain
    
    Dim startPoint As Esprit.Point
    Dim midPoint As Esprit.Point
    Dim endPoint As Esprit.Point
    
    mRefGraphicObject = Document.FeatureRecognition.CreatePartProfileShadow(mSelection, Document.Planes.Item("0DEG"), espFeatureChains)
    For Each ly In Document.Layers
        If (ly.Name = "STL") Then
            mRefGraphicObject(0).Layer = ly
        End If
    Next
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = mRefGraphicObject(0).Layer.Name And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType <> espWorkCoordinate)) Then
            If (.Key > 0) Then
                .Grouped = False
            End If
        End If
        If (.Layer.Name = mRefGraphicObject(0).Layer.Name And (.GraphicObjectType = espFeatureChain)) Then
            If (.Key > 0) Then
                Set returnedFC = graphicObject
            End If
        End If
        End With
    Next
    
    'Call mSelection.Translate(10, 0, 0)
    Dim graphicTemp As Esprit.graphicObject
    Dim lnLine As Esprit.Line
    Dim sgSegment As Esprit.Segment
    Dim sgArc As Esprit.Arc
    
    Dim dLeftEnd As Double
    Dim dRIghtEnd As Double
    Dim i As Integer
    
    dLeftEnd = 0
    dRIghtEnd = 0
    'For Each lnLine In Esprit.Document.Lines
    For i = 1 To returnedFC.Count
        
         If returnedFC.Item(i).TypeName = "Line" Then
            lnLine = returnedFC.Item(i)
            dRIghtEnd = lnLine.x
            dLeftEnd = lnLine.x
         
            'RightEnd
            If lnLine.x > dRIghtEnd Then
                dRIghtEnd = lnLine.x
            End If
            'LeftEnd
            If lnLine.x < dLeftEnd Then
                dLeftEnd = lnLine.x
            End If
         
         ElseIf returnedFC.Item(i).TypeName = "Segment" Then
            Set sgSegment = returnedFC.Item(i)
            'RightEnd
            If sgSegment.XEnd > sgSegment.XStart Then
                If sgSegment.XEnd > dRIghtEnd Then
                    dRIghtEnd = sgSegment.XEnd
                End If
            Else
                If sgSegment.XStart > dRIghtEnd Then
                    dRIghtEnd = sgSegment.XStart
                End If
            End If
            'LeftEnd
            If sgSegment.XStart < sgSegment.XEnd Then
                If sgSegment.XStart < dLeftEnd Then
                    dLeftEnd = sgSegment.XStart
                End If
            Else
                If sgSegment.XEnd < dLeftEnd Then
                    dLeftEnd = sgSegment.XEnd
                End If
            End If
         ElseIf returnedFC.Item(i).TypeName = "Arc" Then
            Set sgArc = returnedFC.Item(i)
            'RightEnd
            If sgArc.Extremity(espExtremityEnd).x > sgArc.Extremity(espExtremityStart).x Then
                If sgArc.Extremity(espExtremityEnd).x > dRIghtEnd Then
                    dRIghtEnd = sgArc.Extremity(espExtremityEnd).x
                End If
            Else
                If sgArc.Extremity(espExtremityStart).x > dRIghtEnd Then
                    dRIghtEnd = sgArc.Extremity(espExtremityStart).x
                End If
            End If
            'LeftEnd
            If sgArc.Extremity(espExtremityStart).x < sgArc.Extremity(espExtremityEnd).x Then
                If sgArc.Extremity(espExtremityStart).x < dLeftEnd Then
                    dLeftEnd = sgArc.Extremity(espExtremityStart).x
                End If
            Else
                If sgArc.Extremity(espExtremityEnd).x < dLeftEnd Then
                    dLeftEnd = sgArc.Extremity(espExtremityEnd).x
                End If
            End If
        
        Else
            MsgBox ("[GetSTLXEND]Not considered type: " & returnedFC.Item(i).TypeName)
        End If
        
    Next
    
    Dim SS As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set SS = .Item("tmpPartProfile")
        On Error GoTo 0
        If SS Is Nothing Then Set SS = .Add("tmpPartProfile")
    End With

    With SS
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(mRefGraphicObject)  'Add the stl_model to the selection object
    End With
    
    Call SS.Delete
    
    'Get X Right End
    'dReturn = returnedFC.BoundingBoxLength + startPoint.X
    'nDirectionCode = 0 : Left / 1 : Right / else : Error
    If nDirectionCode = 0 Then
        dReturn = dLeftEnd
    ElseIf nDirectionCode = 1 Then
        dReturn = dRIghtEnd
    Else
        dReturn = -999 'error
    End If
    'GetCutOffXRightEnd = dCutOffX(3)
    Document.ActiveLayer = lyrObjectInitial
    Document.Refresh
    
EndGetSTLXEnd:
    GetSTLXEnd = dReturn
End Function


Sub GetTurningProfile(Optional ByVal bTolerance As Double = 0.1)
    On Error Resume Next
    On Error GoTo 0
    
    'Tolerance set to 0.1 (default 공차)
    Dim bOriTolerance As Double
    bOriTolerance = Application.Configuration.ConfigurationFeatureRecognition.Tolerance
    If Document.SystemUnit = espInch Then
        Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bTolerance * 3.9
    Else
        Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bTolerance
    End If
        
    Dim layerObject As Esprit.Layer
    Dim graphicObject As Esprit.graphicObject
    Dim stlObject As Esprit.STL_Model
    Dim solidObject As Esprit.Solid
    'Select Group: STL & 경계소재(11)
    For Each layerObject In Esprit.Document.Layers
        layerObject.Visible = False
    Next
    Document.Refresh
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = "STL" And (.GraphicObjectType = espSTL_Model)) Then
            If (.Key > 0) Then
                Set stlObject = graphicObject
                .Grouped = True
                .Layer.Visible = True
            End If
        ElseIf (.Layer.Number = 11 And (.GraphicObjectType = espSolidModel)) Then
            If (.Key > 0) Then
                Set solidObject = graphicObject
                .Grouped = True
                .Layer.Visible = True
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
        
        
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("subGetPartProfile")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("subGetPartProfile")
    End With

    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(stlObject)  'Add the stl_model to the selection object
        Call .Add(solidObject)  'Add the stl_model to the selection object
    End With
    
    Dim mRefGraphicObject() As Esprit.graphicObject
    Dim returnedFC As Esprit.FeatureChain
    
    Dim startPoint As Esprit.Point
    Dim midPoint As Esprit.Point
    Dim endPoint As Esprit.Point
    
    Dim radian As Double
    Dim PI As Double
    PI = 3.14159265
    
    radian = 1 * PI / 180
    
    
    With Document
        .ActiveLayer = .Layers.Item(2) 'Set Turning Profile Layer to Active
        'mRefGraphicObject = .FeatureRecognition.CreateTurningProfile(mSelection, Document.Planes.Item("0DEG"), espTurningProfileOD, espSegmentsArcs, espTurningProfileLocationTop, 0.0001, 0.0001, radian)
        mRefGraphicObject = .FeatureRecognition.CreateTurningProfile(mSelection, Document.Planes.Item("0DEG"), espTurningProfileOD, espSegmentsArcs, espTurningProfileLocationTop, bTolerance, bTolerance, radian)
    End With
    
    Document.Refresh
    'Tolerance back to original value
    Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bOriTolerance

End Sub

Sub GetTurningProfile_EditBoundary(STLRightEndX As Double)
    On Error Resume Next
    On Error GoTo 0
    
        
    Dim layerObject As Esprit.Layer
    Dim graphicObject As Esprit.graphicObject
    Dim stlObject As Esprit.STL_Model
    Dim solidObject As Esprit.Solid
    Dim segmentObject As Esprit.Segment
    Dim segmentSelected As Esprit.Segment
    
    'Select Group: STL & 경계소재(11)
    For Each layerObject In Esprit.Document.Layers
        layerObject.Visible = False
    Next
    Document.Refresh
    
    Dim minX As Double
    minX = 999
    
    For Each segmentObject In Esprit.Document.Segments
        With segmentObject
        If (.Layer.Number = 1 And (Round(.YStart, 5) = Round(.YEnd, 5)) And (.XStart < .XEnd And .XStart < minX) And (.XStart <= STLRightEndX And STLRightEndX < .XEnd)) Then
            If (.Key > 0) Then
                minX = .XStart
                .Grouped = True
                .Layer.Visible = True
                Set segmentSelected = segmentObject
            End If
        ElseIf (.Layer.Number = 1 And (Round(.YStart, 5) = Round(.YEnd, 5)) And (.XStart > .XEnd And .XEnd < minX) And (.XEnd <= STLRightEndX And STLRightEndX < .XStart)) Then
            If (.Key > 0) Then
                minX = .XEnd
                .Grouped = True
                .Layer.Visible = True
                Set segmentSelected = segmentObject
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
        
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("GetTurningProfile_EditBoundary")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("GetTurningProfile_EditBoundary")
    End With

    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(segmentSelected)  'Add the segment to the selection object
    End With
    
    Call mSelection.Translate(0, 1.25, 0, 0)
    
    Dim mRefGraphicObject() As Esprit.graphicObject
    Dim returnedFC As Esprit.FeatureChain
    
    Dim startPoint As Esprit.Point
    Dim midPoint As Esprit.Point
    Dim endPoint As Esprit.Point
    
    Document.Refresh
    
End Sub

Function GetPartProfileSTL(Optional ByVal bTolerance As Double = 0.1) As Esprit.graphicObject()
    On Error Resume Next
    On Error GoTo 0
        
    'Tolerance set to 0.1 (default 공차)
    Dim bOriTolerance As Double
    bOriTolerance = Application.Configuration.ConfigurationFeatureRecognition.Tolerance
    If Document.SystemUnit = espInch Then
        Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bTolerance * 3.9
    Else
        Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bTolerance
    End If
        
        
    Dim layerObject As Esprit.Layer
    Dim graphicObject As Esprit.graphicObject
    Dim stlObject As Esprit.STL_Model
    'Select Group: STL
    For Each layerObject In Esprit.Document.Layers
        layerObject.Visible = False
    Next
    Document.Refresh
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = "STL" And (.GraphicObjectType = espSTL_Model)) Then
            If (.Key > 0) Then
                Set stlObject = graphicObject
                .Grouped = True
                .Layer.Visible = True
                Set layerObject = .Layer
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
        
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("GetPartProfileSTL")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("GetPartProfileSTL")
    End With

    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(stlObject)  'Add the stl_model to the selection object
    End With
    
    Dim mRefGraphicObject() As Esprit.graphicObject
    Dim returnedFC As Esprit.FeatureChain
    
    Dim startPoint As Esprit.Point
    Dim midPoint As Esprit.Point
    Dim endPoint As Esprit.Point
    
    Dim radian As Double
    
    
    With Document
        .ActiveLayer = layerObject 'Set Turning Profile Layer to Active
        GetPartProfileSTL = .FeatureRecognition.CreatePartProfileShadow(mSelection, Document.Planes.Item("0DEG"), espSegmentsArcs)
    End With
    
    Document.Refresh
    'Tolerance back to original value
    Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bOriTolerance
    
End Function

Function TurnSTL(strParse As String) As Integer

    Document.OpenUndoTransaction
    Dim stl As Esprit.STL_Model
    Dim mSelection As Esprit.SelectionSet
    
    TurnSTL = 0
    
    Dim str() As String
    Dim dDegree As Double
    str = Split(strParse, ",")
    If strParse = "" Then
        Call MsgBox("Please set the parameters properly.")
        TurnSTL = -1
        Exit Function
    End If
    
    If Not (str(0) = "X" Or str(0) = "Y" Or str(0) = "Z") Then
        Call MsgBox("First Letter must be X, Y, or Z.")
        TurnSTL = -901
        Exit Function
    End If
    If Not (IsNumeric(str(1))) Then
        Call MsgBox("Second part must be numeric between -180 ~ 180.")
        TurnSTL = -902
        Exit Function
    End If
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = "STL" And (.GraphicObjectType = espSTL_Model)) Then
            If (.Key > 0) Then
                Set stl = graphicObject
            End If
        End If
        End With
    Next

    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("tmpSTL")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("tmpSTL")
    End With

    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(stl)  'Add the stl_model to the selection object
    End With
    
    Call Document.CloseUndoTransaction(True)
    
    
    ' Call mSelection.Translate(5, 0, 0)
    '1. Rotate
    Dim iLine As Esprit.Line
    Dim iPoint As Esprit.Point
    Set iPoint = Document.Points.Add(0, 0, 0)
    
    If str(0) = "X" Then
        Set iLine = Document.Lines.Add(iPoint, 1, 0, 0)
    Else
        If str(0) = "Y" Then
            Set iLine = Document.Lines.Add(iPoint, 0, 1, 0)
        End If
    End If
    
    Dim degree As Double
    degree = CDbl(str(1))
    
    'Rotate by the Axis & Degrees from the parameter
    radian = degree * PI / 180
    Call mSelection.Rotate(iLine, radian, 0)
    
    Document.GraphicsCollection.Remove (iPoint.GraphicsCollectionIndex)
    Document.GraphicsCollection.Remove (iLine.GraphicsCollectionIndex)
    
    Document.Refresh
    TurnSTL = 1
    
End Function


'Step1
Sub CheckSTL(strSTLFilePath As String)
'1. merge STL file in 'STL' Layer
    'Call Document.MergeFile("C:\Users\user\Desktop\기본설정\TEST\Osstem TS,GS Standard ver.1.esp")
    'Document.Refresh
    Dim lyrObjectInitial As Esprit.Layer
    Dim lyrObject As Esprit.Layer
    Set lyrObjectInitial = Document.ActiveLayer
    
    For Each lyrObject In Esprit.Document.Layers
        With lyrObject
        If (.Name = "STL") Then
            Document.ActiveLayer = lyrObject
        End If
        End With
    Next
    
    Dim nCount As Integer
    nCount = 0
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = "STL" And (.GraphicObjectType = espSTL_Model)) Then
            If (.Key > 0) Then
                nCount = nCount + 1
                'Document.GraphicsCollection.Remove (graphicObject.GraphicsCollectionIndex)
                'Call MsgBox("Previous STL model has been deleted.", , "CAM Automation")
                'Document.Refresh
            End If
        End If
        End With
    Next
    
    If nCount = 1 Then
        Exit Sub
    ElseIf nCount > 1 Then
        Call MsgBox("More than 2 STL models are in the document. They is being deleted.", , "CAM Automation")
        For Each graphicObject In Esprit.Document.GraphicsCollection
            With graphicObject
            If (.Layer.Name = "STL" And (.GraphicObjectType = espSTL_Model)) Then
                If (.Key > 0) Then
                    Document.GraphicsCollection.Remove (graphicObject.GraphicsCollectionIndex)
                    'Call MsgBox("Previous STL model has been deleted.", , "CAM Automation")
                End If
            End If
            End With
        Next
        Call Document.MergeFile(strSTLFilePath)
    Else
        Call Document.MergeFile(strSTLFilePath)
    End If
    
    
    Document.Refresh

End Sub

