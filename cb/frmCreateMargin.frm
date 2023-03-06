VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateMargin 
   Caption         =   "[5]DEG 경계소재 & 마진"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3600
   OleObjectBlob   =   "frmCreateMargin.frx":0000
End
Attribute VB_Name = "frmCreateMargin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub chkREAR_Click()
    Dim strWorkPlane As String
    strWorkPlane = ""
    
    If chkREAR = True Then
        If (Document.ActivePlane.Name <> "FACE" And Document.ActivePlane.Name <> "REAR") Then
            Call SaveSetting("frmCreateMargin", "SaveWorkPlane", "SaveWorkPlane", Document.ActivePlane.Name)
        End If
        Document.ActivePlane = Document.Planes("REAR")
    ElseIf chkREAR = False Then
        strWorkPlane = GetSetting("frmCreateMargin", "SaveWorkPlane", "SaveWorkPlane", "0DEG")
        If strWorkPlane = "" Then strWorkPlane = "0DEG"
        
        Select Case strWorkPlane
        Case "0DEG"
            Call cmdENDMILL_000DEG_Click
        Case "90DEG"
            Call cmdENDMILL_090DEG_Click
        Case "180DEG"
            Call cmdENDMILL_180DEG_Click
        Case "270DEG"
            Call cmdENDMILL_270DEG_Click
        Case Else    ' Other values.
            Debug.Print "Not in 0, 90, 180, 270DEG."
            Call cmdENDMILL_000DEG_Click
        End Select
        
    End If
    
    Document.Refresh
End Sub


Private Sub UserForm_Initialize()
   Me.Left = GetSetting("Userform Positioning", "Position-Left-" + Me.Name, "Left", 0)
   Me.Top = GetSetting("Userform Positioning", "Position-Top-" + Me.Name, "Top", 0)
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   Call SaveSetting("Userform Positioning", "Position-Left-" + Me.Name, "Left", Me.Left)
   Call SaveSetting("Userform Positioning", "Position-Top-" + Me.Name, "Top", Me.Top)
End Sub

Private Sub chkShowLatheStock_Click()
    Call Document.Windows.ActiveWindow.SetMask(espViewMaskLatheStock, chkShowLatheStock.Value)
    Document.Refresh
End Sub

Private Sub cmdCreateMargin_Click()
    CreateMargin (Document.ActivePlane.Name)
    UnsuppressOperation (Document.ActivePlane.Name)
    Call cmdReGenerateOp_Click
End Sub

Private Sub cmdENDMILL_000DEG_Click()
    InitializeLayerForMargin ("0DEG")
    lblWorkPlane.Caption = "ODEG"
    lblWorkPlane.Font.Bold = True
    
    cmdENDMILL_000DEG.Font.Bold = True
    cmdENDMILL_090DEG.Font.Bold = False
    cmdENDMILL_180DEG.Font.Bold = False
    cmdENDMILL_270DEG.Font.Bold = False
    
    Call SaveSetting("frmCreateMargin", "SaveWorkPlane", "SaveWorkPlane", Document.ActivePlane.Name)
    chkREAR = False
    
    Call setCurrentTechValues("0DEG", "2")
End Sub

Private Sub cmdENDMILL_090DEG_Click()
    InitializeLayerForMargin ("90DEG")
    lblWorkPlane.Caption = "9ODEG"
    lblWorkPlane.Font.Bold = True

    cmdENDMILL_000DEG.Font.Bold = False
    cmdENDMILL_090DEG.Font.Bold = True
    cmdENDMILL_180DEG.Font.Bold = False
    cmdENDMILL_270DEG.Font.Bold = False

    Call SaveSetting("frmCreateMargin", "SaveWorkPlane", "SaveWorkPlane", Document.ActivePlane.Name)
    chkREAR = False

    Call setCurrentTechValues("90DEG", "4")
End Sub

Private Sub cmdENDMILL_180DEG_Click()
    InitializeLayerForMargin ("180DEG")
    lblWorkPlane.Caption = "180DEG"
    lblWorkPlane.Font.Bold = True
    
    cmdENDMILL_000DEG.Font.Bold = False
    cmdENDMILL_090DEG.Font.Bold = False
    cmdENDMILL_180DEG.Font.Bold = True
    cmdENDMILL_270DEG.Font.Bold = False
    
    Call SaveSetting("frmCreateMargin", "SaveWorkPlane", "SaveWorkPlane", Document.ActivePlane.Name)
    chkREAR = False
    
    Call setCurrentTechValues("180DEG", "6")
End Sub

Private Sub cmdENDMILL_270DEG_Click()
    InitializeLayerForMargin ("270DEG")
    lblWorkPlane.Caption = "270DEG"
    lblWorkPlane.Font.Bold = True

    cmdENDMILL_000DEG.Font.Bold = False
    cmdENDMILL_090DEG.Font.Bold = False
    cmdENDMILL_180DEG.Font.Bold = False
    cmdENDMILL_270DEG.Font.Bold = True

    Call SaveSetting("frmCreateMargin", "SaveWorkPlane", "SaveWorkPlane", Document.ActivePlane.Name)
    chkREAR = False

    Call setCurrentTechValues("270DEG", "8")
End Sub

Private Sub InitializeLayerForMargin(strWorkPlaneName As String)
    Dim ly As Esprit.Layer
    For Each ly In Document.Layers
        If (ly.Name = "STL" Or ly.Name = "기본값" Or ly.Name = strWorkPlaneName + " CROSS BALL ENDMILL" Or ly.Name = strWorkPlaneName + " 마진") Then
            ly.Visible = True
            If (ly.Name = strWorkPlaneName + " 마진") Then
            Document.ActiveLayer = ly
            End If
        Else
            ly.Visible = False
        End If
    Next

    Dim strWorkPlane As String
    strWorkPlane = strWorkPlaneName
    Document.ActivePlane = Document.Planes(strWorkPlane)

    Document.Refresh
End Sub
Function CountSegments(strLayerName As String, Optional bOutputSegmentsInfo As Boolean = True) As Integer
    Dim i As Integer
    Dim nCount As Integer
    
    If bOutputSegmentsInfo Then
        'Application.OutputWindow.Clear
    End If
    
    i = 0
    For Each segmentObject In Esprit.Document.Segments
        With segmentObject
        If (.Layer.Name = strLayerName) Then
            i = i + 1
            If ((Round(.YStart, 5) < Round(.YEnd, 5))) Then
                .Reverse
            End If
            
            If bOutputSegmentsInfo Then
                Application.OutputWindow.Text ("KeySegmentUserMade(" & i & "): " & .Key & vbCrLf)
                Application.OutputWindow.Text ("XStart, XEnd:" & CStr(.XStart) & ", " & CStr(.XEnd) & vbCrLf)
                Application.OutputWindow.Text ("YStart, YEnd:" & CStr(.YStart) & ", " & CStr(.YEnd) & vbCrLf)
                Application.OutputWindow.Text ("ZStart, ZEnd:" & CStr(.ZStart) & ", " & CStr(.ZEnd) & vbCrLf)
            End If
        End If
        End With
    Next
    CountSegments = i
End Function

Function NormalizeSegments(strLayerName As String, Optional bOutputSegmentsInfo As Boolean = True) As Integer
    Dim i As Integer
    Dim nCount As Integer
    
    If bOutputSegmentsInfo Then
        'Application.OutputWindow.Clear
    End If
    
    i = 0
    For Each segmentObject In Esprit.Document.Segments
        With segmentObject
        If (.Layer.Name = strLayerName) Then
            i = i + 1
            .XStart = Round(.ZStart, 5)
            .YStart = Round(.YStart, 5)
            .ZStart = Round(.ZStart, 5)
            .XEnd = Round(.XEnd, 5)
            .YEnd = Round(.YEnd, 5)
            .ZEnd = Round(.ZEnd, 5)
            If ((Round(.YStart, 5) < Round(.YEnd, 5))) Then
                .Reverse
            End If
            
            If bOutputSegmentsInfo Then
                Application.OutputWindow.Text ("KeySegmentUserMade(" & i & "): " & .Key & vbCrLf)
                Application.OutputWindow.Text ("XStart, XEnd:" & CStr(.XStart) & ", " & CStr(.XEnd) & vbCrLf)
                Application.OutputWindow.Text ("YStart, YEnd:" & CStr(.YStart) & ", " & CStr(.YEnd) & vbCrLf)
                Application.OutputWindow.Text ("ZStart, ZEnd:" & CStr(.ZStart) & ", " & CStr(.ZEnd) & vbCrLf)
            End If
        End If
        End With
    Next
    NormalizeSegments = i
End Function


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

Sub PrintSegmentInfo(sgParamSegment As Esprit.Segment)
    Application.OutputWindow.Text ("sgParamSegmentLayer: " & sgParamSegment.Layer.Name & vbCrLf)
    Application.OutputWindow.Text ("sgParamSegmentKey: " & sgParamSegment.Key & vbCrLf)
    Application.OutputWindow.Text ("XStart, XEnd:" & CStr(sgParamSegment.XStart) & ", " & CStr(sgParamSegment.XEnd) & vbCrLf)
    Application.OutputWindow.Text ("YStart, YEnd:" & CStr(sgParamSegment.YStart) & ", " & CStr(sgParamSegment.YEnd) & vbCrLf)
    Application.OutputWindow.Text ("ZStart, ZEnd:" & CStr(sgParamSegment.ZStart) & ", " & CStr(sgParamSegment.ZEnd) & vbCrLf)
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

Private Sub cmdPlane90B_Click()
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
        nDegree = (CInt(strDegree) + 360 - 90) Mod 360
        strDegree = CStr(nDegree) + "DEG"
        
        Document.ActivePlane = Document.Planes(strDegree)
    End If

    Document.Refresh

End Sub

Private Sub UnsuppressOperation(strWorkPlaneName As String)
    'Browse Operations
    Dim Op As Esprit.Operation
    For Each Op In Application.Document.Operations
        If Not (Op.Feature Is Nothing) Then
        'InStr(1, strWorkPlaneName, "DEG", vbTextCompare)
          'If Op.Feature.Name = "8-1. 0DEG CROSS BALL ENDMILL R0.75" Then
          If (InStr(1, Op.Name, " " + strWorkPlaneName, vbTextCompare) <> 0 Or InStr(1, Op.Name, "." + strWorkPlaneName, vbTextCompare) <> 0) And InStr(1, Op.Name, "CROSS BALL ENDMILL R0.75", vbTextCompare) Then
              DoEvents
              Op.Suppress = False
              Op.NeedsReexecute = True
              DoEvents
              Op.Rebuild
              DoEvents
          End If
        End If
    Next

    'release resource
    Set Op = Nothing

End Sub

Private Sub cmdReGenerateOp_Click()
    Dim SelectedDegName As String
    SelectedDegName = getSelectedDEG()
    If SelectedDegName = "" Then
        MsgBox ("Please select 0/90/180/270 DEG button first.")
        Exit Sub
    End If
    
    Dim strOperationOrder As String
    Dim strOperationOrder2 As String
    strOperationOrder = ""
    strOperationOrder2 = ""
    
    If SelectedDegName = "0DEG" Then
        strOperationOrder = "1"
        strOperationOrder2 = "2"
    ElseIf SelectedDegName = "90DEG" Then
        strOperationOrder = "3"
        strOperationOrder2 = "4"
    ElseIf SelectedDegName = "180DEG" Then
        strOperationOrder = "5"
        strOperationOrder2 = "6"
    ElseIf SelectedDegName = "270DEG" Then
        strOperationOrder = "7"
        strOperationOrder2 = "8"
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''
'' re-create Toolpath 8-2/4/6/8. 0DEG-1 CROSS BALL ENDMILL R0.75
    Dim SelectedOpName As String
    Dim SelectedOpName2 As String
    SelectedOpName = "8-" + strOperationOrder + ". " + getSelectedDEG + " CROSS BALL ENDMILL R0.75"
    'SelectedOpName = "NOT IN USE"
    SelectedOpName2 = "8-" + strOperationOrder2 + ". " + getSelectedDEG + "-1" + " CROSS BALL ENDMILL R0.75"
    
    Dim tech As EspritTechnology.Technology
    Dim techTLMPP As EspritTechnology.TechLatheMoldParallelPlanes
    'espTechLatheMillContour1
    'techTLMPP.StepOver
    'Dim dBottomZLimit As Double
    'dBottomZLimit = Conversion.CDbl(txtBottomZLimit.Text)
    
    Dim dStepOver As Double
    dStepOver = Conversion.CDbl(txtStepOver.Text)
    
    Dim Op As Esprit.Operation
    For Each Op In Application.Document.Operations
        If (Op.Name = SelectedOpName) Then
            Set tech = Op.Technology
            Set techTLMPP = tech
            techTLMPP.BoundaryProfiles = getBoundaryProfiles(SelectedDegName, True)
            'techTLMPP.BottomZLimit = dBottomZLimit
        ElseIf (Op.Name = SelectedOpName2) Then
            Set tech = Op.Technology
            Set techTLMPP = tech
            techTLMPP.BoundaryProfiles = getBoundaryProfiles(SelectedDegName, False)
            techTLMPP.StepOver = dStepOver
        End If
    Next
    
    Set techTLMPP = Nothing
    Set tech = Nothing
    Set Op = Nothing
    
End Sub

Private Function getBoundaryProfiles(strWorkPlane As String, Optional bTakeFirstKeyOnly As Boolean = False) As String
    Dim strTempBoundaryProfiles As String
    strTempBoundaryProfiles = ""
    
    Dim goRef As Esprit.graphicObject
    Dim lyMargin As String
    Dim lyCrossBallEndmill As String
    Dim strTempLayer As String
    
    Dim strKeyTemp As String
    
    lyMargin = strWorkPlane + " 마진"
    lyCrossBallEndmill = strWorkPlane + " CROSS BALL ENDMILL"

    For Each goRef In Esprit.Document.GraphicsCollection
        With goRef
        If ((.Layer.Name = lyMargin Or .Layer.Name = lyCrossBallEndmill) And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType = espFeatureChain)) Then
            If (.Key > 0) Then
                If strTempBoundaryProfiles = "" Then
                    strTempBoundaryProfiles = CStr(espFeatureChain) + "," + CStr(.Key)
                    strKeyTemp = .Key
                ElseIf bTakeFirstKeyOnly = True And CInt(.Key) > CInt(strKeyTemp) Then
                    strTempBoundaryProfiles = CStr(espFeatureChain) + "," + CStr(.Key)
                    strKeyTemp = .Key
                ElseIf bTakeFirstKeyOnly = False Then
                    strTempBoundaryProfiles = strTempBoundaryProfiles + "|" + CStr(espFeatureChain) + "," + CStr(.Key)
                End If
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
    
    Set goRef = Nothing
    
    getBoundaryProfiles = strTempBoundaryProfiles
    
End Function

Private Function getSelectedDEG() As String
    Dim strSelectedPlaneName As String
    If cmdENDMILL_000DEG.Font.Bold = True Then
        strSelectedPlaneName = "0DEG"
    ElseIf cmdENDMILL_090DEG.Font.Bold = True Then
        strSelectedPlaneName = "90DEG"
    ElseIf cmdENDMILL_180DEG.Font.Bold = True Then
        strSelectedPlaneName = "180DEG"
    ElseIf cmdENDMILL_270DEG.Font.Bold = True Then
        strSelectedPlaneName = "270DEG"
    Else
        strSelectedPlaneName = ""
    End If
    
    getSelectedDEG = strSelectedPlaneName
End Function


Private Sub setCurrentTechValues(strWorkPlaneName As String, strOperationOrder As String)
''''''''''''''''''''''''''''''''''''''''''''''''''
'' re-create Toolpath (3_ROUGH_ENDMILL)
    Dim SelectedOpName As String
    Dim tech As EspritTechnology.Technology
    Dim techTLMPP As EspritTechnology.TechLatheMoldParallelPlanes
    Dim Op As Esprit.Operation
    
    Select Case strOperationOrder
'    Case 1, 3, 5, 7
'        SelectedOpName = "8-" + strOperationOrder + ". " + getSelectedDEG + " CROSS BALL ENDMILL R0.75"
'        For Each Op In Application.Document.Operations
'            If Op.Name = SelectedOpName Then
'                Set tech = Op.Technology
'                If tech.TechnologyType = espTechLatheMoldParallelPlanes Then
'                    Set techTLMPP = tech
'                    txtBottomZLimit.Text = CStr(techTLMPP.BottomZLimit)
'                End If
'            End If
'        Next
        
    Case 2, 4, 6, 8
        SelectedOpName = "8-" + strOperationOrder + ". " + getSelectedDEG + "-1" + " CROSS BALL ENDMILL R0.75"
        For Each Op In Application.Document.Operations
            If Op.Name = SelectedOpName Then
                Set tech = Op.Technology
                If tech.TechnologyType = espTechLatheMoldParallelPlanes Then
                    Set techTLMPP = tech
                    txtStepOver.Text = CStr(techTLMPP.StepOver)
                End If
            End If
        Next
    Case Else    ' Other values.
        Debug.Print "Not in 2, 4, 6, 8."
        Exit Sub
    End Select
    
    Set techTLMPP = Nothing
    Set tech = Nothing
    Set Op = Nothing
    
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' DEG 경계소재

Private Sub cmdDEGBorder_000DEG_Click()
    InitializeLayerForBorder ("0DEG")
    lblDEGBorder_WorkPlane.Caption = "ODEG"
    lblDEGBorder_WorkPlane.Font.Bold = True
    
    cmdDEGBorder_000DEG.Font.Bold = True
    cmdDEGBorder_090DEG.Font.Bold = False
    cmdDEGBorder_180DEG.Font.Bold = False
    cmdDEGBorder_270DEG.Font.Bold = False
    
    Call SaveSetting("frmCreateMargin_pageDEGBorder", "chkDEGBorder_SaveWorkPlane", "chkDEGBorder_SaveWorkPlane", Document.ActivePlane.Name)
    Call setDEGBorder_CurrentTechValues("0DEG", "1")

End Sub
Private Sub cmdDEGBorder_090DEG_Click()
    InitializeLayerForBorder ("90DEG")
    lblDEGBorder_WorkPlane.Caption = "9ODEG"
    lblDEGBorder_WorkPlane.Font.Bold = True
    
    cmdDEGBorder_000DEG.Font.Bold = False
    cmdDEGBorder_090DEG.Font.Bold = True
    cmdDEGBorder_180DEG.Font.Bold = False
    cmdDEGBorder_270DEG.Font.Bold = False
    
    
    Call SaveSetting("frmCreateMargin_pageDEGBorder", "chkDEGBorder_SaveWorkPlane", "chkDEGBorder_SaveWorkPlane", Document.ActivePlane.Name)
    chkDEGBorder_FACE = False
    chkDEGBorder_REAR = False
    Call setDEGBorder_CurrentTechValues("90DEG", "3")

End Sub
Private Sub cmdDEGBorder_180DEG_Click()
    InitializeLayerForBorder ("180DEG")
    lblDEGBorder_WorkPlane.Caption = "180DEG"
    lblDEGBorder_WorkPlane.Font.Bold = True
    
    cmdDEGBorder_000DEG.Font.Bold = False
    cmdDEGBorder_090DEG.Font.Bold = False
    cmdDEGBorder_180DEG.Font.Bold = True
    cmdDEGBorder_270DEG.Font.Bold = False
    
    Call SaveSetting("frmCreateMargin_pageDEGBorder", "chkDEGBorder_SaveWorkPlane", "chkDEGBorder_SaveWorkPlane", Document.ActivePlane.Name)
    chkDEGBorder_FACE = False
    chkDEGBorder_REAR = False
    Call setDEGBorder_CurrentTechValues("180DEG", "5")

End Sub
Private Sub cmdDEGBorder_270DEG_Click()
    InitializeLayerForBorder ("270DEG")
    lblDEGBorder_WorkPlane.Caption = "270DEG"
    lblDEGBorder_WorkPlane.Font.Bold = True
    
    cmdDEGBorder_000DEG.Font.Bold = False
    cmdDEGBorder_090DEG.Font.Bold = False
    cmdDEGBorder_180DEG.Font.Bold = False
    cmdDEGBorder_270DEG.Font.Bold = True
    
    Call SaveSetting("frmCreateMargin_pageDEGBorder", "chkDEGBorder_SaveWorkPlane", "chkDEGBorder_SaveWorkPlane", Document.ActivePlane.Name)
    Call setDEGBorder_CurrentTechValues("270DEG", "7")

End Sub

Private Sub InitializeLayerForBorder(strWorkPlaneName As String)
    Dim ly As Esprit.Layer
    For Each ly In Document.Layers
        If (ly.Name = "STL" Or ly.Name = "기본값" Or ly.Name = strWorkPlaneName + " 경계소재") Then
            ly.Visible = True
            If (ly.Name = strWorkPlaneName + " 경계소재") Then
            Document.ActiveLayer = ly
            End If
        Else
            ly.Visible = False
        End If
    Next

    Dim strWorkPlane As String
    strWorkPlane = strWorkPlaneName
    Document.ActivePlane = Document.Planes(strWorkPlane)

    Document.Refresh
End Sub

Private Sub setDEGBorder_CurrentTechValues(strWorkPlaneName As String, strOperationOrder As String)
''''''''''''''''''''''''''''''''''''''''''''''''''
'' re-create Toolpath (3_ROUGH_ENDMILL)
    Dim SelectedOpName As String
    Dim tech As EspritTechnology.Technology
    Dim techTLMPP As EspritTechnology.TechLatheMoldParallelPlanes
    Dim Op As Esprit.Operation
    
    Select Case strOperationOrder
    Case 1, 3, 5, 7
        SelectedOpName = "8-" + strOperationOrder + ". " + getsetDEGBorder_SelectedDEG + " CROSS BALL ENDMILL R0.75"
        For Each Op In Application.Document.Operations
            If Op.Name = SelectedOpName Then
                Set tech = Op.Technology
                If tech.TechnologyType = espTechLatheMoldParallelPlanes Then
                    Set techTLMPP = tech
                    txtDEGBorder_BottomZLimit.Text = CStr(techTLMPP.BottomZLimit)
                End If
            End If
        Next
    Case Else    ' Other values.
        Debug.Print "Not in 1, 3, 5, 7."
        Exit Sub
    End Select
    
    Set techTLMPP = Nothing
    Set tech = Nothing
    Set Op = Nothing
    
End Sub
Private Function getsetDEGBorder_SelectedDEG() As String
    Dim strSelectedPlaneName As String
    If cmdDEGBorder_000DEG.Font.Bold = True Then
        strSelectedPlaneName = "0DEG"
    ElseIf cmdDEGBorder_090DEG.Font.Bold = True Then
        strSelectedPlaneName = "90DEG"
    ElseIf cmdDEGBorder_180DEG.Font.Bold = True Then
        strSelectedPlaneName = "180DEG"
    ElseIf cmdDEGBorder_270DEG.Font.Bold = True Then
        strSelectedPlaneName = "270DEG"
    Else
        strSelectedPlaneName = ""
    End If
    
    getsetDEGBorder_SelectedDEG = strSelectedPlaneName
End Function


Private Sub cmdDEGBorder_Plane90A_Click()
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
Private Sub cmdDEGBorder_Plane90B_Click()
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
        nDegree = (CInt(strDegree) + 360 - 90) Mod 360
        strDegree = CStr(nDegree) + "DEG"
        
        Document.ActivePlane = Document.Planes(strDegree)
    End If

    Document.Refresh
End Sub

Private Sub chkDEGBorder_FACE_Click()
    Dim strWorkPlane As String
    strWorkPlane = ""
    
    If chkDEGBorder_FACE = True Then
        If chkDEGBorder_REAR.Value = True Then
            chkDEGBorder_REAR.Value = False
        End If
        
        If (Document.ActivePlane.Name <> "FACE" And Document.ActivePlane.Name <> "REAR") Then
            Call SaveSetting("frmCreateMargin_pageDEGBorder", "chkDEGBorder_SaveWorkPlane", "chkDEGBorder_SaveWorkPlane", Document.ActivePlane.Name)
        End If
        Document.ActivePlane = Document.Planes("FACE")
    ElseIf chkDEGBorder_FACE = False Then
        If chkDEGBorder_REAR.Value = False Then
            strWorkPlane = GetSetting("frmCreateMargin_pageDEGBorder", "chkDEGBorder_SaveWorkPlane", "chkDEGBorder_SaveWorkPlane", "0DEG")
            If strWorkPlane = "" Then strWorkPlane = "0DEG"
            
            Select Case strWorkPlane
            Case "0DEG"
                Call cmdDEGBorder_000DEG_Click
            Case "90DEG"
                Call cmdDEGBorder_090DEG_Click
            Case "180DEG"
                Call cmdDEGBorder_180DEG_Click
            Case "270DEG"
                Call cmdDEGBorder_270DEG_Click
            Case Else    ' Other values.
                Debug.Print "Not in 0, 90, 180, 270DEG."
                Call cmdDEGBorder_000DEG_Click
            End Select
            
            'Document.ActivePlane = Document.Planes(strWorkPlane)
        End If
    End If
    
    Document.Refresh
End Sub

Private Sub chkDEGBorder_REAR_Click()
    Dim strWorkPlane As String
    strWorkPlane = ""
    
    If chkDEGBorder_REAR = True Then
        If chkDEGBorder_FACE.Value = True Then
            chkDEGBorder_FACE.Value = False
        End If
        If (Document.ActivePlane.Name <> "FACE" And Document.ActivePlane.Name <> "REAR") Then
            Call SaveSetting("frmCreateMargin_pageDEGBorder", "chkDEGBorder_SaveWorkPlane", "chkDEGBorder_SaveWorkPlane", Document.ActivePlane.Name)
        End If
        Document.ActivePlane = Document.Planes("REAR")
    ElseIf chkDEGBorder_REAR = False Then
        If chkDEGBorder_FACE.Value = False Then
            strWorkPlane = GetSetting("frmCreateMargin_pageDEGBorder", "chkDEGBorder_SaveWorkPlane", "chkDEGBorder_SaveWorkPlane", "0DEG")
            If strWorkPlane = "" Then strWorkPlane = "0DEG"
            
            Select Case strWorkPlane
            Case "0DEG"
                Call cmdDEGBorder_000DEG_Click
            Case "90DEG"
                Call cmdDEGBorder_090DEG_Click
            Case "180DEG"
                Call cmdDEGBorder_180DEG_Click
            Case "270DEG"
                Call cmdDEGBorder_270DEG_Click
            Case Else    ' Other values.
                Debug.Print "Not in 0, 90, 180, 270DEG."
                Call cmdDEGBorder_000DEG_Click
            End Select
            
            'Document.ActivePlane = Document.Planes(strWorkPlane)
        End If
    End If
    
    Document.Refresh
End Sub


Private Sub cmdDEGBorder_Regenerate_Click()
    Dim SelectedDegName As String
    SelectedDegName = getDEGBorder_SelectedDEG()
    If SelectedDegName = "" Then
        MsgBox ("Please select 0/90/180/270 DEG button first.")
        Exit Sub
    End If
    
    Dim strOperationOrder As String
    strOperationOrder = ""
    
    If SelectedDegName = "0DEG" Then
        strOperationOrder = "1"
    ElseIf SelectedDegName = "90DEG" Then
        strOperationOrder = "3"
    ElseIf SelectedDegName = "180DEG" Then
        strOperationOrder = "5"
    ElseIf SelectedDegName = "270DEG" Then
        strOperationOrder = "7"
    End If
    
    Call FreeFormCheckElementSolidRefresh(Document.ActiveLayer.Name)
    
''''''''''''''''''''''''''''''''''''''''''''''''''
'' re-create Toolpath 8-1/3/5/7. CROSS BALL ENDMILL R0.75
    Dim SelectedOpName As String
    SelectedOpName = "8-" + strOperationOrder + ". " + SelectedDegName + " CROSS BALL ENDMILL R0.75"
    
    Dim tech As EspritTechnology.Technology
    Dim techTLMPP As EspritTechnology.TechLatheMoldParallelPlanes
    Dim dBottomZLimit As Double
    dBottomZLimit = Conversion.CDbl(txtDEGBorder_BottomZLimit.Text)
    
    Dim Op As Esprit.Operation
    For Each Op In Application.Document.Operations
        If (Op.Name = SelectedOpName) Then
            Set tech = Op.Technology
            Set techTLMPP = tech
            techTLMPP.BottomZLimit = dBottomZLimit
        End If
    Next
    
    Set techTLMPP = Nothing
    Set tech = Nothing
    Set Op = Nothing

End Sub

Private Function getDEGBorder_SelectedDEG() As String
    Dim strSelectedPlaneName As String
    If cmdDEGBorder_000DEG.Font.Bold = True Then
        strSelectedPlaneName = "0DEG"
    ElseIf cmdDEGBorder_090DEG.Font.Bold = True Then
        strSelectedPlaneName = "90DEG"
    ElseIf cmdDEGBorder_180DEG.Font.Bold = True Then
        strSelectedPlaneName = "180DEG"
    ElseIf cmdDEGBorder_270DEG.Font.Bold = True Then
        strSelectedPlaneName = "270DEG"
    Else
        strSelectedPlaneName = ""
    End If
    
    getDEGBorder_SelectedDEG = strSelectedPlaneName
End Function


