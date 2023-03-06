VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPGMText 
   Caption         =   "Engraving Program Number Text"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3255
   OleObjectBlob   =   "frmPGMText.frx":0000
End
Attribute VB_Name = "frmPGMText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Private Sub cmdCreateProjectionText_Click()
    Dim strWorkPlane As String
    strWorkPlane = Document.ActivePlane.Name
    Dim strBaseLayerName As String
    strBaseLayerName = "TEXT"
    
    Call createProjectionFinishing(strWorkPlane, strBaseLayerName)
    Call ReorderOperation
End Sub

Sub createProjectionFinishing(strWorkPlaneName As String, strBaseLayerName As String)
    
    Dim Fc As FeatureChain
    Set Fc = findTextFeatureChain
    
    Dim mSelection As Esprit.SelectionSet
    Set mSelection = findTextSelectionSet
    
    If mSelection.Count = 0 Then
        MsgBox ("Cannot find the Text. Please generate it first.")
        Exit Sub
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''
'' make Toolpath (Projection Finishing)
    
    Dim FileName As String
    FileName = "TEXT_Projection_FInishing.prc"

    Dim M_TechnologyUtility As EspritTechnology.TechnologyUtility
    Set M_TechnologyUtility = Document.TechnologyUtility
    Dim tech() As EspritTechnology.Technology
    tech = M_TechnologyUtility.OpenProcess(EspritUserFolder & FileName)
    
    Dim techLM3DNew As EspritTechnology.TechLatheMill3DProject
    Set techLM3DNew = tech(0)
    techLM3DNew.ProjectElement = CStr(espFeatureChain) + "," + CStr(Fc.Key)
    
    'Rebuild Operations
    Dim fff As Esprit.FreeFormFeature
    Dim fSelected As Esprit.FreeFormFeature
    Dim graphicObject As Esprit.graphicObject
    Dim goSelected As Esprit.graphicObject
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
            If (.Layer.Name = "STL" And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType = espSTL_Model)) Then
                If (.Key > 0) Then
                    Set goSelected = graphicObject
                End If
            End If
        End With
    Next
    
    Document.ActiveWorkCoordinate = Document.WorkCoordinates("GSE3010(CROSS) or GSE3207")
    
    Set fff = Document.FreeFormFeatures.Add()
    Call fff.Add(goSelected, espFreeFormPartSurfaceItem)
    fff.Name = "7 프리폼[TEXT] New"
    
    
    Dim Op As Esprit.Operation
    Set Op = Document.Operations.Add(tech(0), fff)
    Op.Name = "9. " & strWorkPlaneName & " TEXT-" & GetProgramNumber() & " CROSS BALL ENDMILL R0.2"
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''
'' re-create Toolpath 9. xxxDEG 9 CROSS BALL END MILL R0.2
' "9. [XXX]DEG 9 CROSS BALL END MILL R0.2"
'    Dim SelectedOpName As String
'    SelectedOpName = " CROSS BALL ENDMILL R0.2"
'
'    Dim tech As EspritTechnology.Technology
'    Dim techLM3D As EspritTechnology.TechLatheMill3DProject
'
'    Dim Op As Esprit.Operation
'    For Each Op In Application.Document.Operations
'        If InStr(Op.Name, SelectedOpName) > 0 Then
'            Set tech = Op.Technology
'            Set techLM3D = tech
'            techLM3D.ProjectElement = CStr(espFeatureChain) + "," + CStr(Fc.Key)
'            Op.Name = "9. " & strWorkPlane & " TEXT(PGMNo):" & GetProgramNumber() & " CROSS BALL ENDMILL R0.2"
'
'        End If
'    Next
    
    Set techLM3DNew = Nothing
    Set tech(0) = Nothing
    Set Op = Nothing
    
'    For Each fff In Application.Document.FreeFormFeatures
'        Call ChangeFreeFormPartElement(fff, goSelected)
'    Next
    

    Set fff = Nothing
    Set fSelected = Nothing
    Set graphicObject = Nothing
    Set goSelected = Nothing
    
    Dim wkTemp As Esprit.WorkCoordinate
    For Each wkTemp In Esprit.Document.WorkCoordinates
        With wkTemp
        If (InStr(.Name, "XYZ") > 0) Then
            Document.ActiveWorkCoordinate = wkTemp
        End If
        End With
    Next
    Document.Refresh
    
End Sub



Private Sub cmdLeft_Click()
    Dim strWorkPlane As String
    strWorkPlane = Document.ActivePlane.Name
    
    Dim chnText As FeatureChain
    Dim chnTextStartPoint As Point
    Set chnText = findTextFeatureChain
    Set chnTextStartPoint = chnText.Extremity(espExtremityStart)
    
    Dim mSelection As Esprit.SelectionSet
    Set mSelection = findTextSelectionSet
    
    If mSelection.Count = 0 Then
        MsgBox ("Cannot find the Text. Please generate it first.")
        Exit Sub
    End If
    
    Dim dUnit As Double
    dUnit = Conversion.CDbl(txtLeftRight.Text)
    'Move Left
    Call mSelection.Translate(dUnit * (-1), 0, 0)
End Sub
Private Sub cmdToZeroDegree_Click()
    ChangeWorkPlane ("0DEG")
End Sub

Private Sub cmdNext90Degree_Click()
    movePlaneNext (udfPlaneDegree90)
End Sub

Private Sub cmdNext10Dgree_Click()
    movePlaneNext (udfPlaneDegree10)
End Sub
Private Sub cmdBefore10Dgree_Click()
    movePlaneBefore (udfPlaneDegree10)
End Sub

Private Sub cmdPlane90A_Click()
    Dim nPosition As Integer
    Dim strCurrPlane As String
    Dim strDegree As String
    Dim nDegree As Integer
    strCurrPlane = Document.ActivePlane.Name
    nPosition = InStr(1, strCurrPlane, "DEG", vbTextCompare)
    
    If (nPosition = 0) Then
        Document.ActivePlane = Document.Planes("0DEG")
    Else
        strDegree = Replace(strCurrPlane, "DEG", "")
        nDegree = (CInt(strDegree) + 90) Mod 360
        strDegree = CStr(nDegree) + "DEG"
        
        Document.ActivePlane = Document.Planes(strDegree)
    End If

    Document.Refresh
End Sub



Private Sub cmdRight_Click()
    Dim strWorkPlane As String
    strWorkPlane = Document.ActivePlane.Name
    
    Dim chnText As FeatureChain
    Dim chnTextStartPoint As Point
    Set chnText = findTextFeatureChain
    Set chnTextStartPoint = chnText.Extremity(espExtremityStart)
    
    Dim mSelection As Esprit.SelectionSet
    Set mSelection = findTextSelectionSet
    
    If mSelection.Count = 0 Then
        MsgBox ("Cannot find the Text. Please generate it first.")
        Exit Sub
    End If
    
    Dim dUnit As Double
    dUnit = Conversion.CDbl(txtLeftRight.Text)
    'Move Right
    Call mSelection.Translate(dUnit * (1), 0, 0)
End Sub

Private Sub cmdCreateTextFeature_Click()
'1. Set ActiveLayer to "Text"
    setLayersForText
    Document.Refresh

'2. Save Current WorkPlaneName
    Dim strWorkPlane As String
    strWorkPlane = Document.ActivePlane.Name

    Dim strmSelectionIndex As String
    strmSelectionIndex = strWorkPlane + "PGMText"
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item(strmSelectionIndex)
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add(strmSelectionIndex)
    End With
    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
    End With
    
    Dim engEngraving As ESPRITEngrave.Engraving
    Set engEngraving = New ESPRITEngrave.Engraving
    engEngraving.Text = GetProgramNumber()
    
    Dim EP As Esprit.Point
    MsgBox ("Pick a point where to generate the Text.")
    Set EP = Document.GetPoint("Pick Part Number Location")
    Call mSelection.Add(EP)
    Call mSelection.AlignPlane(Document.Planes(strWorkPlane), espAlignPlaneToGlobalXYZ)
    
    Call EP.SetXyz(EP.x, EP.y, EP.Z + 10)
    Call mSelection.AlignPlane(Document.Planes(strWorkPlane), espAlignPlaneFromGlobalXYZ)
    
'    engEngraving.CreateType = espEngraveCreateFeatures
'    engEngraving.TypeOfPattern = espEngravePatternLinear
'    engEngraving.CreateQuantity = espEngraveCreateSingle
'    engEngraving.AngleOfPattern = 0
'    engEngraving.FontType = espEngraveFontWindows
'    engEngraving.Font.Name = "Arial"
'    engEngraving.Font.Size = 6
'    engEngraving.HeightSpacing = 20
'    engEngraving.WidthSpacing = 20
'    engEngraving.Clearance
'    engEngraving.TotalDepth = 0.2
'    engEngraving.EnableDoubleBack = True
    
    ' compare the Planes.Count before and after to remove any extra planes created
    Dim originalPlanesCount As Long
    originalPlanesCount = Document.Planes.Count
    
    Call engEngraving.Execute(EP)
    lblTextWorkPlane.Caption = strWorkPlane
    With Document.Planes
        If .Count > originalPlanesCount Then
            For i = .Count To (originalPlanesCount + 1) Step -1
                Debug.Print "Removing " & .Item(i).Name
                Call .Remove(i)
            Next
        End If
    End With
    
    Document.Refresh
    
    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
    End With
    
    Set EP = Nothing
    Set engEngraving = Nothing
    Set mSelection = Nothing
    
End Sub


Private Sub cmdFindTextObject_Click()
'1. Check TextWorkPlane
'    Dim strTextWorkPlane As String
'    strTextWorkPlane = lblTextWorkPlane.Caption
'
'    If strTextWorkPlane = "" Then
'        MsgBox ("Please Generate Text First.")
'        Return
'    End If

'2.
    If (lblTextWorkPlane.Caption <> "") Then
        Document.ActivePlane = Document.Planes(lblTextWorkPlane.Caption)
    End If
    
    If findTextSelectionSet.Count = 0 Then
        MsgBox ("Cannot find the Text. Please generate it first.")
    End If
    
    Document.Refresh

End Sub
Private Function findTextSelectionSet() As SelectionSet
    Dim strmSelectionIndex As String
    strmSelectionIndex = strWorkPlane + "PGMText"
    Dim mSelection As Esprit.SelectionSet
    
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item(strmSelectionIndex)
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add(strmSelectionIndex)
    End With
    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
    End With
    
    Dim goRef As Esprit.graphicObject
    Dim plRef As Esprit.Plane
    Dim lyOri As Esprit.Layer
    Dim strTempLayer As String
    
    setLayersForText
    Set lyOri = Document.ActiveLayer

    For Each goRef In Esprit.Document.GraphicsCollection
        With goRef
        If (.Layer.Name = lyOri.Name And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType = espFeatureChain)) Then
            If (.Key > 0) Then
                .Grouped = True
                Call mSelection.Add(goRef)
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
    
    Set findTextSelectionSet = mSelection
    
End Function
Private Function findTextFeatureChain() As FeatureChain
    Dim chnText As FeatureChain
    Dim rtnText As FeatureChain
    
    setLayersForText
    Set lyOri = Document.ActiveLayer

    For Each chnText In Esprit.Document.FeatureChains
        If (chnText.Layer.Name = lyOri.Name) Then
            Set rtnText = chnText
        End If
    Next
    
    Set findTextFeatureChain = rtnText
End Function



Private Sub ChangeWorkPlane(strWorkPlaneName As String)
    Dim strWorkPlane As String
    strWorkPlane = strWorkPlaneName
    Document.ActivePlane = Document.Planes(strWorkPlane)

    Document.Refresh
End Sub

Private Sub cmdUpDonwMinus_Click()
    Dim strWorkPlane As String
    strWorkPlane = Document.ActivePlane.Name
    
    Dim chnText As FeatureChain
    Dim chnTextStartPoint As Point
    Set chnText = findTextFeatureChain
    Set chnTextStartPoint = chnText.Extremity(espExtremityStart)
    
    Dim mSelection As Esprit.SelectionSet
    Set mSelection = findTextSelectionSet
    
    If mSelection.Count = 0 Then
        MsgBox ("Cannot find the Text. Please generate it first.")
        Exit Sub
    End If
    
    Call mSelection.AlignPlane(Document.Planes(strWorkPlane), espAlignPlaneToGlobalXYZ)

    
    Dim dUnit As Double
    dUnit = Conversion.CDbl(txtUpDown.Text)
    
    Dim nWorkPlaneInInt As Integer
    nWorkPlaneInInt = getWorkPlaneToInt(Replace(Document.ActivePlane.Name, "DEG", ""))
    
    'Move Down
    If Not (nWorkPlaneInInt >= 0 And nWorkPlaneInInt < 360) Then
        MsgBox ("Please check your Workplane in 0DEG to 350DEG")
        Exit Sub
    'ElseIf (nWorkPlaneInInt \ 180 > 0) Then
    '    Call mSelection.Translate(0, -dUnit * (-1), 0)
    Else
        Call mSelection.Translate(0, -dUnit, 0)
    End If
    
    Call mSelection.AlignPlane(Document.Planes(strWorkPlane), espAlignPlaneFromGlobalXYZ)
    
End Sub

Private Sub cmdUpDonwPlus_Click()
    Dim strWorkPlane As String
    strWorkPlane = Document.ActivePlane.Name
    
    Dim chnText As FeatureChain
    Dim chnTextStartPoint As Point
    Set chnText = findTextFeatureChain
    Set chnTextStartPoint = chnText.Extremity(espExtremityStart)
    
    Dim mSelection As Esprit.SelectionSet
    Set mSelection = findTextSelectionSet
    
    If mSelection.Count = 0 Then
        MsgBox ("Cannot find the Text. Please generate it first.")
        Exit Sub
    End If
    
    Call mSelection.AlignPlane(Document.Planes(strWorkPlane), espAlignPlaneToGlobalXYZ)
    'Document.ActivePlane = Document.Planes("0DEG") 'Must be 0DEG for align to XYZ
    
    
    Dim dUnit As Double
    dUnit = Conversion.CDbl(txtUpDown.Text)
    
    Dim nWorkPlaneInInt As Integer
    nWorkPlaneInInt = getWorkPlaneToInt(Replace(Document.ActivePlane.Name, "DEG", ""))
    
    
    
    If Not (nWorkPlaneInInt >= 0 And nWorkPlaneInInt < 360) Then
        MsgBox ("Please check your Workplane in 0DEG to 350DEG")
        Exit Sub
    'ElseIf (nWorkPlaneInInt \ 180 > 0) Then
    '    Call mSelection.Translate(0, dUnit * (-1), 0)
    Else
        Call mSelection.Translate(0, dUnit, 0)
    End If
    
    Call mSelection.AlignPlane(Document.Planes(strWorkPlane), espAlignPlaneFromGlobalXYZ)

End Sub

Public Function LocalX(P As Esprit.Point, Optional LocalPlane As Esprit.Plane) As Double

    If LocalPlane Is Nothing Then

        Set LocalPlane = Document.ActivePlane

    End If

    LocalX = (P.x - LocalPlane.x) * LocalPlane.Ux

    LocalX = LocalX + (P.y - LocalPlane.y) * LocalPlane.Uy

    LocalX = LocalX + (P.Z - LocalPlane.Z) * LocalPlane.Uz

End Function
Public Function LocalY(P As Esprit.Point, Optional LocalPlane As Esprit.Plane) As Double

    If LocalPlane Is Nothing Then

        Set LocalPlane = Document.ActivePlane

    End If

    LocalY = (P.x - LocalPlane.x) * LocalPlane.Vx

    LocalY = LocalY + (P.y - LocalPlane.y) * LocalPlane.Vy

    LocalY = LocalY + (P.Z - LocalPlane.Z) * LocalPlane.Vz

End Function

 

Public Function LocalZ(P As Esprit.Point, Optional LocalPlane As Esprit.Plane) As Double

    If LocalPlane Is Nothing Then

        Set LocalPlane = Document.ActivePlane

    End If

    LocalZ = (P.x - LocalPlane.x) * LocalPlane.Wx

    LocalZ = LocalZ + (P.y - LocalPlane.y) * LocalPlane.Wy

    LocalZ = LocalZ + (P.Z - LocalPlane.Z) * LocalPlane.Wz

End Function
Private Sub cmdRotatePlus_Click()
    Dim strWorkPlane As String
    strWorkPlane = Document.ActivePlane.Name
    
    Dim chnText As FeatureChain
    Dim chnTextStartPoint As Point
    Set chnText = findTextFeatureChain
    Set chnTextStartPoint = chnText.Extremity(espExtremityStart)
    
    Dim mSelection As Esprit.SelectionSet
    Set mSelection = findTextSelectionSet
    
    If mSelection.Count = 0 Then
        MsgBox ("Cannot find the Text. Please generate it first.")
        Exit Sub
    End If
    
    'Rotate
    Dim iLine As Esprit.Line
    'Dim iPoint As Esprit.Point
    'Set iPoint = Document.Points.Add(0, 0, 0)
    Set iLine = getTheOriginAxis()
    
    Dim dUnit As Double
    dUnit = Conversion.CDbl(txtRotate.Text)
    'Rotate by the Axis & Degrees from the parameter
    radian = dUnit * PI / 180
    Call mSelection.Rotate(iLine, radian, 0)
    
    'Document.GraphicsCollection.Remove (iPoint.GraphicsCollectionIndex)
    'Document.GraphicsCollection.Remove (iLine.GraphicsCollectionIndex)
    
    Document.Refresh
End Sub
Private Sub cmdRotateMinus_Click()
    Dim strWorkPlane As String
    strWorkPlane = Document.ActivePlane.Name
    
    Dim chnText As FeatureChain
    Dim chnTextStartPoint As Point
    Set chnText = findTextFeatureChain
    Set chnTextStartPoint = chnText.Extremity(espExtremityStart)
    
    Dim mSelection As Esprit.SelectionSet
    Set mSelection = findTextSelectionSet
    
    If mSelection.Count = 0 Then
        MsgBox ("Cannot find the Text. Please generate it first.")
        Exit Sub
    End If
    
    'Rotate Reverse
    Dim iLine As Esprit.Line
    Set iLine = getTheOriginAxis()
    
    Dim dUnit As Double
    dUnit = Conversion.CDbl(txtRotate.Text)
    dUnit = (360 - dUnit) Mod 360
    'Rotate by the Axis & Degrees from the parameter
    radian = dUnit * PI / 180
    Call mSelection.Rotate(iLine, radian, 0)
    
    Document.Refresh
End Sub

Sub CreateMargin(strWorkPlaneName As String)
    On Error Resume Next
    On Error GoTo 0
    
'    Application.OutputWindow.Clear

    
'1> Check Status
'1)  Layer Check: 0DEG 마진, 90DEG 마진, 180DEG 마진, 270DEG 마진
    With Document.ActiveLayer
    If Not (.Name = "0DEG 마진" Or .Name = "90DEG 마진" Or .Name = "180DEG 마진" Or .Name = "270DEG 마진") Then
        Call MsgBox("As an aactive layer, should select a layer in (0DEG 마진, 90DEG 마진, 180DEG 마진, 270DEG 마진)", vbOKOnly, "Alert")
        Exit Sub
    End If
    End With

'2) Segments check: at least have 2 segments in the layer
    If CountSegments(strWorkPlaneName + " 마진") < 2 Then
        Call MsgBox("You should make more than 2 segments at least.", vbOKOnly, "Alert")
        Exit Sub
    End If



'
' create the feature chains
'

'2> Get mSelection for segments in the layer
'1) Set Save Original Layer As lyOri
    Dim lyOri As Esprit.Layer
    Dim strOriginalLayer As String
    strOriginalLayer = Document.ActiveLayer.Name
    
    For Each ly In Document.Layers
        If (ly.Name = strOriginalLayer) Then
            Set lyOri = ly
        End If
    Next
    
    Document.ActiveLayer = lyOri
    Document.ActiveLayer.Visible = True
    'Document.Refresh

'2)  0DEG 선택(기본)
    Dim strWorkPlane As String
    strWorkPlane = strWorkPlaneName
    Document.ActivePlane = Document.Planes(strWorkPlane)
    'Document.Refresh
    
'3)
    Dim strmSelectionIndex As String
    strmSelectionIndex = strWorkPlane + "CreditMargin"
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item(strmSelectionIndex)
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add(strmSelectionIndex)
    End With
    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
    End With
    
    Dim goRef As Esprit.graphicObject
    Dim plRef As Esprit.Plane
    Dim lyTemp As Esprit.Layer
    Dim strTempLayer As String
    
    strTempLayer = "TempLayer"
    For Each ly In Document.Layers
        If (ly.Name = strTempLayer) Then
            Call Document.Layers.Remove(strTempLayer)
        End If
    Next
    Set lyTemp = Document.Layers.Add(strTempLayer)
    Document.ActiveLayer = lyTemp

    For Each goRef In Esprit.Document.GraphicsCollection
        With goRef
        If (.Layer.Name = lyOri.Name And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType = espSegment)) Then
            If (.Key > 0) Then
                .Grouped = True
                Call mSelection.Add(goRef)
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next

    lyTemp.Visible = True
    Call mSelection.AlignPlane(Document.Planes(strWorkPlane), espAlignPlaneToGlobalXYZ)
    Document.ActivePlane = Document.Planes("0DEG") 'Must be 0DEG for align to XYZ
    
    Call mSelection.ChangeLayer(lyTemp, 0)
    Document.Refresh

'
'4) Make margin automatically
    Dim layerObject As Esprit.Layer
    Dim graphicObject As Esprit.graphicObject
    Dim stlObject As Esprit.STL_Model
    Dim solidObject As Esprit.Solid
    Dim segmentObject As Esprit.Segment
    Dim segmentSelected As Esprit.Segment

    Dim nKeySegmentUserMade(1000) As String

    For Each segmentObject In Esprit.Document.Segments
        With segmentObject
        If (.Layer.Name = lyTemp.Name) Then
            i = i + 1
            If ((Round(.YStart, 5) < Round(.YEnd, 5))) Then
                .Reverse
            End If
            nKeySegmentUserMade(i) = .Key
        End If
        End With
    Next
    nCount = i

    Dim dMin As Double
    dMin = 999
    Dim dMax As Double
    dMax = -999

    'nSegmentA: Top Part Segment
    Dim nSegmentA As Integer
    Dim sgSegmentA As Esprit.Segment
    'nSegmentA: Bottom Part Segment
    Dim nSegmentB As Integer
    Dim sgSegmentB As Esprit.Segment

    Dim sgSegmentC As Esprit.Segment
    Dim sgSegmentD As Esprit.Segment
    Dim sgSegmentE As Esprit.Segment

    
    'Get the SegmentA
    For i = 1 To nCount
        If (Document.Segments.Item(nKeySegmentUserMade(i)).YStart > dMax) Then
            dMax = Document.Segments.Item(nKeySegmentUserMade(i)).YStart
            nSegmentA = i
        End If
    Next i
    Set sgSegmentA = Document.Segments.Item(nKeySegmentUserMade(nSegmentA))

    'Get the SegmentB
    For i = 1 To nCount
        If (Document.Segments.Item(nKeySegmentUserMade(i)).YEnd < dMin) Then
            dMin = Document.Segments.Item(nKeySegmentUserMade(i)).YEnd
            nSegmentB = i
        End If
    Next i
    Set sgSegmentB = Document.Segments.Item(nKeySegmentUserMade(nSegmentB))

    Call PrintSegmentInfo(sgSegmentA)
    Call PrintSegmentInfo(sgSegmentB)

    Dim nXBase As Double
    nXBase = -0.5
    
    Set sgSegmentC = Document.Segments.Add(Document.GetPoint(nXBase, sgSegmentA.YStart, 0), Document.GetPoint(nXBase, sgSegmentB.YEnd, 0))
    Set sgSegmentD = Document.Segments.Add(Document.GetPoint(nXBase, sgSegmentA.YStart, 0), Document.GetPoint(sgSegmentA.XStart, sgSegmentA.YStart, 0))
    Set sgSegmentE = Document.Segments.Add(Document.GetPoint(nXBase, sgSegmentB.YEnd, 0), Document.GetPoint(sgSegmentB.XEnd, sgSegmentB.YEnd, 0))
    Call mSelection.Add(sgSegmentC)
    Call mSelection.Add(sgSegmentD)
    Call mSelection.Add(sgSegmentE)

    Call mSelection.AlignPlane(Document.Planes(strWorkPlane), espAlignPlaneFromGlobalXYZ)
    Document.ActivePlane = Document.Planes(strWorkPlane)

    Document.ActiveLayer = lyOri
    Call mSelection.ChangeLayer(Document.ActiveLayer, 0)
    
    Dim GraphicObj() As Esprit.graphicObject
    If mSelection.Count > 0 Then
        GraphicObj = Document.FeatureRecognition.CreateAutoChains(mSelection)
    End If

    'DeleteSmallPointChainFeature
    Call DeleteSmallPointChainFeature(lyOri.Name, 0.2)
    Document.Refresh
    
    'Count FC
    Dim nCnt As Integer
    nCnt = 0
    For Each fcCnt In Document.FeatureChains
        If (fcCnt.Layer.Name = lyOri.Name) Then
            nCnt = nCnt + 1
        End If
    Next
    

    For Each ly In Document.Layers
        If (ly.Name = strTempLayer) Then
            Call Document.Layers.Remove(strTempLayer)
        End If
    Next
    
End Sub
Private Sub CommandButton3_Click()
    Dim P1 As Esprit.Point
    Dim P2 As Esprit.Point
    Set P1 = Document.Points.Add(0, 0, 1)
    Set P2 = Document.Points.Add(1, 1, 0)
    
    Dim radian As Double
    Dim PI As Double
    PI = 3.14159265
    radian = 1 * PI / 180
    
    Document.Windows.ActiveWindow.Rotate radian, radian, radian
    Document.Windows.ActiveWindow.Refresh
End Sub

Private Sub CommandButton4_Click()
    Dim radian As Double
    Dim PI As Double
    PI = 3.14159265
    radian = 1 * PI / 180

    Document.Windows.ActiveWindow.Rotate -radian, -radian, -radian
    Document.Windows.ActiveWindow.Refresh
End Sub

Private Sub setLayersForText()

    Dim ly As Esprit.Layer
    For Each ly In Document.Layers
        If (ly.Name = "STL" Or ly.Name = "기본값") Then
            ly.Visible = True
        ElseIf (ly.Name = "TEXT") Then
            ly.Visible = True
            Document.ActiveLayer = ly
        Else
            ly.Visible = False
        End If
    Next
    Document.Refresh

End Sub
Private Sub UserForm_Initialize()
    '
    ' first get default values from registry
    ' note that last argument is the default if registry key does not exist
    '
'    Dim ScaleFactor As Double
'    If Document.SystemUnit = espMetric Then
'        ScaleFactor = 25.4
'    Else
'        ScaleFactor = 1
'    End If

   Me.Left = GetSetting("Userform Positioning", "Position-Left-" + Me.Name, "Left", 0)
   Me.Top = GetSetting("Userform Positioning", "Position-Top-" + Me.Name, "Top", 0)


    txtUpDown.Text = GetSetting("CAMAutomationFrmPGMText", "Recent", "txtUpDown", 1)
    txtLeftRight.Text = GetSetting("CAMAutomationFrmPGMText", "Recent", "txtLeftRight", 0.1)
    txtRotate.Text = GetSetting("CAMAutomationFrmPGMText", "Recent", "txtRotate", 5)

End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    '
    ' now record the values back to the registry (in case the user changed them)
    '
    Call SaveSetting("CAMAutomationFrmPGMText", "Recent", "txtUpDown", txtUpDown.Text)
    Call SaveSetting("CAMAutomationFrmPGMText", "Recent", "txtLeftRight", txtLeftRight.Text)
    Call SaveSetting("CAMAutomationFrmPGMText", "Recent", "txtRotate", txtRotate.Text)
    
   Call SaveSetting("Userform Positioning", "Position-Left-" + Me.Name, "Left", Me.Left)
   Call SaveSetting("Userform Positioning", "Position-Top-" + Me.Name, "Top", Me.Top)
    
End Sub

Function getWorkPlaneToInt(strWorkPlane As String) As Integer
    Dim nDegree As Integer
    nDegree = -1
    nDegree = Conversion.CInt(strWorkPlane)
    
    getWorkPlaneToInt = nDegree
End Function
