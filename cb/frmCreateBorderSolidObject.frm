VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateBorderSolidObject 
   Caption         =   "Check Rough ML & Create Border Solid"
   ClientHeight    =   4305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3480
   OleObjectBlob   =   "frmCreateBorderSolidObject.frx":0000
End
Attribute VB_Name = "frmCreateBorderSolidObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False














Private Sub cmdCreateSolidObject_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Esprit Not Supported
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'1. Check Arcs in the layer
'1) more than 3 Arcs
'2) Check all of the arcs are connected
'3) Make a selection with all of the arcs

'2. Move the Selection to X,-0.5
'3. 보스/돌출잘라내기
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' TAB: 경계소재-2
Private Sub cmdBefore10DgreeBorder2_Click()
    setLayersFor ("경계소재-2")
    movePlaneBefore (udfPlaneDegree10)
End Sub

Private Sub cmdBefore90DegreeBorder2_Click()
    setLayersFor ("경계소재-2")
    movePlaneBefore (udfPlaneDegree90)
End Sub
Private Sub cmdNext10DgreeBorder2_Click()
    setLayersFor ("경계소재-2")
    movePlaneNext (udfPlaneDegree10)
End Sub

Private Sub cmdNext90DegreeBorder2_Click()
    setLayersFor ("경계소재-2")
    movePlaneNext (udfPlaneDegree90)
End Sub

Private Sub cmdPlane0DegreeBorder2_Click()
    setLayersFor ("경계소재-2")
    Document.ActivePlane = Document.Planes("0DEG")
End Sub

Private Sub cmdLeftBorder2_Click()
    Dim strWorkPlane As String
    strWorkPlane = Document.ActivePlane.Name
    
    Dim mSelection As Esprit.SelectionSet
    Set mSelection = findBorder2SelectionSet(espSolidModel)
    
    If mSelection.Count = 0 Then
        MsgBox ("Cannot find the Solid in [" + Document.ActiveLayer.Name + "]. Please generate it first.")
        Exit Sub
    End If
    
    Dim dUnit As Double
    dUnit = Conversion.CDbl(txtLeftRightBorder2.Text)
    'Move Left
    Call mSelection.Translate(dUnit * (-1), 0, 0)
End Sub


Private Sub cmdRightBorder2_Click()
    Dim strWorkPlane As String
    strWorkPlane = Document.ActivePlane.Name
    
    Dim mSelection As Esprit.SelectionSet
    Set mSelection = findBorder2SelectionSet(espSolidModel)
    
    If mSelection.Count = 0 Then
        MsgBox ("Cannot find the Solid in [" + Document.ActiveLayer.Name + "]. Please generate it first.")
        Exit Sub
    End If
    
    Dim dUnit As Double
    dUnit = Conversion.CDbl(txtLeftRightBorder2.Text)
    'Move Right
    Call mSelection.Translate(dUnit * (1), 0, 0)
End Sub

Private Sub cmdRotatePlusBorder2_Click()
    Dim strWorkPlane As String
    strWorkPlane = Document.ActivePlane.Name
    
    Dim mSelection As Esprit.SelectionSet
    Set mSelection = findBorder2SelectionSet(espSolidModel)
    
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
    dUnit = Conversion.CDbl(txtRotateBorder2.Text)
    'Rotate by the Axis & Degrees from the parameter
    radian = dUnit * PI / 180
    Call mSelection.Rotate(iLine, radian, 0)
    
    Document.Refresh
End Sub
Private Sub cmdRotateMinusBorder2_Click()
    Dim strWorkPlane As String
    strWorkPlane = Document.ActivePlane.Name
    
    Dim mSelection As Esprit.SelectionSet
    Set mSelection = findBorder2SelectionSet(espSolidModel)
    
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
    dUnit = Conversion.CDbl(txtRotateBorder2.Text)
    dUnit = (360 - dUnit) Mod 360
    'Rotate by the Axis & Degrees from the parameter
    radian = dUnit * PI / 180
    Call mSelection.Rotate(iLine, radian, 0)
    
    Document.Refresh
End Sub



Private Function findBorder2GraphicObject(Optional ByVal paramEspGraphicObjectType As EspritConstants.espGraphicObjectType = espSolidModel) As Esprit.graphicObject
    Dim chnBorder3 As Esprit.graphicObject
    Dim rtnBorder3 As Esprit.graphicObject
    
    setLayersFor ("경계소재-2")
    Set lyOri = Document.ActiveLayer

    For Each chnBorder2 In Esprit.Document.GraphicsCollection
        If (chnBorder2.Layer.Name = lyOri.Name) And (chnBorder2.GraphicObjectType = paramEspGraphicObjectType) Then
            Set rtnBorder2 = chnBorder2
        End If
    Next
    
    Set findBorder2GraphicObject = rtnBorder2
End Function

Private Function findBorder2SelectionSet(Optional ByVal paramEspGraphicObjectType As EspritConstants.espGraphicObjectType = espSolidModel) As SelectionSet
    Dim strmSelectionIndex As String
    strmSelectionIndex = strWorkPlane + "Border2"
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
    
    setLayersFor ("경계소재-2")
    Set lyOri = Document.ActiveLayer

    For Each goRef In Esprit.Document.GraphicsCollection
        With goRef
        If (.Layer.Name = lyOri.Name And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType = paramEspGraphicObjectType)) Then
            If (.Key > 0) Then
                .Grouped = True
                Call mSelection.Add(goRef)
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
    
    Set findBorder2SelectionSet = mSelection
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' TAB: 경계소재-3
Private Sub cmdPlane0DegreeBorder3_Click()
    setLayersFor ("경계소재-3")
    Document.ActivePlane = Document.Planes("0DEG")
End Sub


Private Sub cmdPlaneFaceBorder3_Click()
    setLayersFor ("경계소재-3")
    Document.ActivePlane = Document.Planes("FACE")
End Sub

Private Sub setLayersFor(strBaseLayerName As String)

    Dim ly As Esprit.Layer
    For Each ly In Document.Layers
        If (ly.Name = "STL" Or ly.Name = "기본값") Then
            ly.Visible = True
        ElseIf (ly.Name = strBaseLayerName) Then
            ly.Visible = True
            Document.ActiveLayer = ly
        Else
            ly.Visible = False
        End If
    Next
    Document.Refresh
End Sub

Private Sub cmdLeft_Click()
    Dim strWorkPlane As String
    strWorkPlane = Document.ActivePlane.Name
    
    Dim mSelection As Esprit.SelectionSet
    Set mSelection = findBorder3SelectionSet(espSolidModel)
    
    If mSelection.Count = 0 Then
        MsgBox ("Cannot find the Solid in [" + Document.ActiveLayer.Name + "]. Please generate it first.")
        Exit Sub
    End If
    
    Dim dUnit As Double
    dUnit = Conversion.CDbl(txtLeftRight.Text)
    'Move Left
    Call mSelection.Translate(dUnit * (-1), 0, 0)
End Sub


Private Sub cmdRight_Click()
    Dim strWorkPlane As String
    strWorkPlane = Document.ActivePlane.Name
    
    Dim mSelection As Esprit.SelectionSet
    Set mSelection = findBorder3SelectionSet(espSolidModel)
    
    If mSelection.Count = 0 Then
        MsgBox ("Cannot find the Solid in [" + Document.ActiveLayer.Name + "]. Please generate it first.")
        Exit Sub
    End If
    
    Dim dUnit As Double
    dUnit = Conversion.CDbl(txtLeftRight.Text)
    'Move Right
    Call mSelection.Translate(dUnit * (1), 0, 0)
End Sub

Private Sub cmdRotatePlus_Click()
    Dim strWorkPlane As String
    strWorkPlane = Document.ActivePlane.Name
    
    Dim mSelection As Esprit.SelectionSet
    Set mSelection = findBorder3SelectionSet(espSolidModel)
    
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
    
    Dim mSelection As Esprit.SelectionSet
    Set mSelection = findBorder3SelectionSet(espSolidModel)
    
    If mSelection.Count = 0 Then
        MsgBox ("Cannot find the Text. Please generate it first.")
        Exit Sub
    End If
    
    'Rotate
    Dim iLine As Esprit.Line
    Dim iPoint As Esprit.Point
    'Dim iPoint As Esprit.Point
    'Set iPoint = Document.Points.Add(0, 0, 0)
    Set iLine = getTheOriginAxis()
    
    Dim dUnit As Double
    dUnit = Conversion.CDbl(txtRotate.Text)
    dUnit = (360 - dUnit) Mod 360
    'Rotate by the Axis & Degrees from the parameter
    radian = dUnit * PI / 180
    Call mSelection.Rotate(iLine, radian, 0)
    
    'Document.GraphicsCollection.Remove (iPoint.GraphicsCollectionIndex)
    'Document.GraphicsCollection.Remove (iLine.GraphicsCollectionIndex)
    
    Document.Refresh
End Sub

Private Function findBorder3GraphicObject(Optional ByVal paramEspGraphicObjectType As EspritConstants.espGraphicObjectType = espSolidModel) As Esprit.graphicObject
    Dim chnBorder3 As Esprit.graphicObject
    Dim rtnBorder3 As Esprit.graphicObject
    
    setLayersFor ("경계소재-3")
    Set lyOri = Document.ActiveLayer

    For Each chnBorder3 In Esprit.Document.GraphicsCollection
        If (chnBorder3.Layer.Name = lyOri.Name) And (chnBorder3.GraphicObjectType = paramEspGraphicObjectType) Then
            Set rtnBorder3 = chnBorder3
        End If
    Next
    
    Set findBorder3GraphicObject = rtnBorder3
End Function

Private Function findBorder3SelectionSet(Optional ByVal paramEspGraphicObjectType As EspritConstants.espGraphicObjectType = espSolidModel) As SelectionSet
    Dim strmSelectionIndex As String
    strmSelectionIndex = strWorkPlane + "Border3"
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
    
    setLayersFor ("경계소재-3")
    Set lyOri = Document.ActiveLayer

    For Each goRef In Esprit.Document.GraphicsCollection
        With goRef
        If (.Layer.Name = lyOri.Name And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType = paramEspGraphicObjectType)) Then
            If (.Key > 0) Then
                .Grouped = True
                Call mSelection.Add(goRef)
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
    
    Set findBorder3SelectionSet = mSelection
    
End Function


Private Sub cmdSetBasicLocateSolidModel_Click()
    Dim mSelection As Esprit.SelectionSet
    Set mSelection = findBorder3SelectionSet(espSolidModel)
    Document.ActivePlane = Document.Planes("0DEG")
    
    If mSelection.Count = 0 Then
        MsgBox ("Cannot find the Text. Please generate it first.")
        Exit Sub
    End If
    
    'Move Left
    Call mSelection.Translate(0.5 * (-1), 0, 0)
End Sub

Private Sub CommandButton1_Click()

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

    txtDistanceBorder2.Text = GetSetting("CAMAutomationFrmCreateBorderSolidObject", "Recent", "txtDistanceBorder2", 10)
    txtLeftRightBorder2.Text = GetSetting("CAMAutomationFrmCreateBorderSolidObject", "Recent", "txtLeftRightBorder2", 0.5)
    txtRotateBorder2.Text = GetSetting("CAMAutomationFrmCreateBorderSolidObject", "Recent", "txtRotateBorder2", 1)

    txtDistanceBorder3.Text = GetSetting("CAMAutomationFrmCreateBorderSolidObject", "Recent", "txtDistanceBorder3", 3)
    txtLeftRight.Text = GetSetting("CAMAutomationFrmCreateBorderSolidObject", "Recent", "txtLeftRight", 0.5)
    txtRotate.Text = GetSetting("CAMAutomationFrmCreateBorderSolidObject", "Recent", "txtRotate", 1)
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    '
    ' now record the values back to the registry (in case the user changed them)
    '
   
    Call SaveSetting("CAMAutomationFrmCreateBorderSolidObject", "Recent", "txtDistanceBorder2", txtDistanceBorder2.Text)
    Call SaveSetting("CAMAutomationFrmCreateBorderSolidObject", "Recent", "txtLeftRightBorder2", txtLeftRightBorder2.Text)
    Call SaveSetting("CAMAutomationFrmCreateBorderSolidObject", "Recent", "txtRotateBorder2", txtRotateBorder2.Text)

    Call SaveSetting("CAMAutomationFrmCreateBorderSolidObject", "Recent", "txtDistanceBorder3", txtDistanceBorder3.Text)
    Call SaveSetting("CAMAutomationFrmCreateBorderSolidObject", "Recent", "txtLeftRight", txtLeftRight.Text)
    Call SaveSetting("CAMAutomationFrmCreateBorderSolidObject", "Recent", "txtRotate", txtRotate.Text)
    
    Call SaveSetting("Userform Positioning", "Position-Left-" + Me.Name, "Left", Me.Left)
    Call SaveSetting("Userform Positioning", "Position-Top-" + Me.Name, "Top", Me.Top)

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkShowLatheStock_Click()
    Call Document.Windows.ActiveWindow.SetMask(espViewMaskLatheStock, chkShowLatheStock.Value)
    Document.Refresh
End Sub

Private Sub cmdCreateRoughEndMill_Click()
    Dim SelectedDegName As String
    SelectedDegName = getSelectedDEG()
    If SelectedDegName = "" Then
        MsgBox ("Please select 0/120/240 DEG button first.")
        Exit Sub
    End If
    
    Dim strOperationOrder As String
    strOperationOrder = ""
    
    If SelectedDegName = "0DEG" Then
        strOperationOrder = "1"
    ElseIf SelectedDegName = "120DEG" Then
        strOperationOrder = "2"
    ElseIf SelectedDegName = "240DEG" Then
        strOperationOrder = "3"
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''
'' re-create Toolpath (3_ROUGH_ENDMILL)
    Dim SelectedOpName As String
    SelectedOpName = "3-" + strOperationOrder + ". ROUGH_ENDMILL_" + getSelectedDEG
    
'    Dim M_TechnologyUtility As EspritTechnology.TechnologyUtility
'    Set M_TechnologyUtility = Document.TechnologyUtility
'    Dim Tech() As EspritTechnology.Technology
'    Tech = M_TechnologyUtility.OpenProcess(EspritUserFolder & FileName)
    
    Dim tech As EspritTechnology.Technology
    Dim techTLMCtr1 As EspritTechnology.TechLatheMillContour1
    'espTechLatheMillContour1
    Dim lRoughPasses As Long
    lRoughPasses = Conversion.CLng(txtNumberOfRoughingPath.Text)
    Dim dTotalDepth As Double
    dTotalDepth = Conversion.CDbl(txtTotalDepth.Text)
    Dim dStartingDepth As Double
    dStartingDepth = Conversion.CDbl(txtStartingDepth.Text)
    
    'txtTotalDepth
    'txtStartingDepth
    
    Dim Op As Esprit.Operation
    For Each Op In Application.Document.Operations
        If Op.Name = SelectedOpName Then
            Set tech = Op.Technology
            Set techTLMCtr1 = tech
            techTLMCtr1.RoughPasses = lRoughPasses
            techTLMCtr1.TotalDepth = dTotalDepth
            techTLMCtr1.StartingDepth = dStartingDepth
            'tech = techTLMCtr1
            'Op.Technology = tech
'              DoEvents
'              Op.NeedsReexecute = True
'              DoEvents
'              Op.Rebuild
'              DoEvents
        End If
    Next

End Sub

Private Sub cmdPlane90A_Click()
    movePlaneNext (udfPlaneDegree90)
End Sub


Private Sub cmdPlane90B_Click()
    movePlaneBefore (udfPlaneDegree90)
End Sub

Private Function getSelectedDEG() As String
    Dim strSelectedPlaneName As String
    If cmdRoughEndMill_000DEG.Font.Bold = True Then
        strSelectedPlaneName = "0DEG"
    ElseIf cmdRoughEndMill_120DEG.Font.Bold = True Then
        strSelectedPlaneName = "120DEG"
    ElseIf cmdRoughEndMill_240DEG.Font.Bold = True Then
        strSelectedPlaneName = "240DEG"
    Else
        strSelectedPlaneName = ""
    End If
    
    getSelectedDEG = strSelectedPlaneName
End Function
Public Sub initialRun()
    Call cmdRoughEndMill_000DEG_Click
End Sub


Private Sub cmdRoughEndMill_000DEG_Click()
    InitializeLayerForMargin ("0DEG")
    Call setCurrentTechValues("0DEG", "1")
    
    
    cmdRoughEndMill_000DEG.Font.Bold = True
    cmdRoughEndMill_120DEG.Font.Bold = False
    cmdRoughEndMill_240DEG.Font.Bold = False
End Sub
Private Sub cmdRoughEndMill_120DEG_Click()
    InitializeLayerForMargin ("120DEG")
    Call setCurrentTechValues("120DEG", "2")
    
    cmdRoughEndMill_000DEG.Font.Bold = False
    cmdRoughEndMill_120DEG.Font.Bold = True
    cmdRoughEndMill_240DEG.Font.Bold = False
    
End Sub

Private Sub cmdRoughEndMill_240DEG_Click()
    InitializeLayerForMargin ("240DEG")
    Call setCurrentTechValues("240DEG", "3")
    
    cmdRoughEndMill_000DEG.Font.Bold = False
    cmdRoughEndMill_120DEG.Font.Bold = False
    cmdRoughEndMill_240DEG.Font.Bold = True
End Sub

Private Sub InitializeLayerForMargin(strWorkPlaneName As String)
    Dim ly As Esprit.Layer
    
    Dim strSetRoughLayer As String
    strSetRoughLayer = "[" + strWorkPlaneName + "]" + "ROUGH ENDMILL R6.0"
    
    For Each ly In Document.Layers
        If (ly.Name = "STL" Or ly.Name = "기본값" Or ly.Name = strSetRoughLayer) Then
            ly.Visible = True
            If (ly.Name = strSetRoughLayer) Then
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

Private Sub setCurrentTechValues(strWorkPlaneName As String, strOperationOrder As String)
''''''''''''''''''''''''''''''''''''''''''''''''''
'' re-create Toolpath (3_ROUGH_ENDMILL)
    Dim SelectedOpName As String
    SelectedOpName = "3-" + strOperationOrder + ". ROUGH_ENDMILL_" + strWorkPlaneName

    
    Dim tech As EspritTechnology.Technology
    Dim techTLMCtr1 As EspritTechnology.TechLatheMillContour1
    
    Dim Op As Esprit.Operation
    For Each Op In Application.Document.Operations
        If Op.Name = SelectedOpName Then
            Set tech = Op.Technology
            Set techTLMCtr1 = tech
            
            txtNumberOfRoughingPath.Text = CStr(techTLMCtr1.RoughPasses)
            txtTotalDepth.Text = CStr(techTLMCtr1.TotalDepth)
            txtStartingDepth.Text = CStr(techTLMCtr1.StartingDepth)
            
        End If
    Next
    
    Set techTLMCtr1 = Nothing
    Set tech = Nothing
    Set Op = Nothing

End Sub

