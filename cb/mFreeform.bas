Attribute VB_Name = "mFreeform"
'Attribute VB_Name = "FreeformFeatureSolidAssociation"
Option Explicit

Private Sub OutputFreeformFeatureSolidList()
    'loop through freeform features
    'output a list of the solids that are associated to the FFF in the output window
    Dim i As Integer
    'use oftype in vb.net
    For i = 1 To Document.Group.Count
        Dim go As Esprit.graphicObject, fff As Esprit.FreeFormFeature, SolidList As String
        On Error Resume Next
        Set go = Document.Group.Item(i)
        Set fff = Nothing
        SolidList = ""
        On Error GoTo 0
        If Not go Is Nothing Then
            If go.GraphicObjectType = espFreeFormFeature Then
                Set fff = go
            End If
            If Not fff Is Nothing Then
                SolidList = GetFreeformSolidIds(fff)
                OutputWindow.Text fff.Name & " - " & SolidList & vbNewLine
            End If
        End If
    Next
    OutputWindow.Visible = True
End Sub

Private Function GetFreeformSolidIds(fff As FreeFormFeature) As String
    Dim i As Integer, J As Integer, fffItem As EspritFeatures.ComFreeFormFeatureElement
    
    ReDim ItemList(1 To fff.Count) As String
    
    'loop through items in FFF, compile a list of bodies (solids, surfaces) that make up that solid
    For i = 1 To fff.Count
        On Error Resume Next
        Set fffItem = fff.Item(i)
        Dim go As graphicObject, sb As EspritSolids.SolidBody, sf As EspritSolids.SolidFace
        Set go = Nothing
        If IsGraphicObj(fffItem.Object) Then
            Set go = fffItem.Object
        Else
            On Error Resume Next
            Set sb = fffItem.Object
            Set sf = fffItem.Object
            On Error GoTo 0
            If Not sb Is Nothing Then
                Set go = GetGraphicObjectFromSolidBody(sb)
            End If
            If Not sf Is Nothing Then
                Set go = GetGraphicObjectFromSolidFace(sf)
            End If
        End If
        
        Dim GoId As String
        
        'get graphic object solid id and add it to the list (if it doesnt exist already)
        If Not go Is Nothing Then
            GoId = go.TypeName & go.Key & "(" & go.Layer.Name & ")"
            J = J + 1
            If Not ArrayContains(GoId, ItemList) Then
                ItemList(J) = GoId
            End If
        End If
        On Error GoTo 0
    Next
    GetFreeformSolidIds = GetSimpleStringFromItemListArray(ItemList)
End Function

Private Function GetGraphicObjectFromSolidBody(ByVal sb As EspritSolids.SolidBody) As graphicObject
    Set GetGraphicObjectFromSolidBody = Nothing
    If sb Is Nothing Then Exit Function
    Dim i As Integer
    For i = 1 To Document.Solids.Count
        If Document.Solids.Item(i).SolidBody.Tag = sb.Tag Then
            Set GetGraphicObjectFromSolidBody = Document.Solids.Item(i)
            Exit Function
        End If
    Next
End Function

Private Function GetGraphicObjectFromSolidFace(ByVal sf As EspritSolids.SolidFace) As graphicObject
    Set GetGraphicObjectFromSolidFace = Nothing
    If sf Is Nothing Then Exit Function
    Dim sb As EspritSolids.SolidBody
    Set sb = sf.SolidBody
    If sb Is Nothing Then Exit Function
    Dim i As Integer
    For i = 1 To Document.Solids.Count
        If Document.Solids.Item(i).SolidBody.Tag = sb.Tag Then
            Set GetGraphicObjectFromSolidFace = Document.Solids.Item(i)
            Exit Function
        End If
    Next
End Function

Private Function ArrayContains(Item As String, ItemArr() As String) As Boolean
    ArrayContains = False
    Dim list() As String, i As Integer
    list = ItemArr
    For i = LBound(list) To UBound(list)
        If list(i) = Item Then
            ArrayContains = True
            Exit Function
        End If
    Next
End Function

Private Function GetSimpleStringFromItemListArray(ItemList() As String) As String
    Dim i As Integer, ret As String
    For i = LBound(ItemList) To UBound(ItemList)
        If ItemList(i) <> "" Then
            If ret = "" Then
                ret = ItemList(i)
            Else
                ret = ret & ", " & ItemList(i)
            End If
        End If
    Next
    GetSimpleStringFromItemListArray = ret
End Function

Private Function IsGraphicObj(TestObject As Object) As Boolean
    Dim TestResult As Esprit.graphicObject
    On Error Resume Next
    Set TestResult = TestObject ' attempt to typecast
    IsGraphicObj = Not (TestResult Is Nothing)
    On Error GoTo 0
End Function
Public Sub RebuildFreeformWithNewSTLAsAPartElement()

    'Rebuild Operations
    Dim fff As Esprit.FreeFormFeature
    Dim fSelected As Esprit.FreeFormFeature
    Dim graphicObject As Esprit.graphicObject
    Dim goSelected As Esprit.graphicObject
    Dim goSelected_ExtraLayer As Esprit.graphicObject
    
    Dim tmpLayer As Esprit.Layer
    
    For Each tmpLayer In Document.Layers
        If (tmpLayer.Name = "STL") Then
            For Each graphicObject In tmpLayer.Document.GraphicsCollection
                With graphicObject
                    If (.Layer.Name = "STL" And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType = espSTL_Model)) Then
                        If (.Key > 0) Then
                            Set goSelected = graphicObject
                        End If
                    End If
                End With
            Next
        End If
        If (tmpLayer.Name = "SPECIAL") Then
            For Each graphicObject In Esprit.Document.GraphicsCollection
                With graphicObject
                    If (.Layer.Name = "SPECIAL" And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType = espSTL_Model)) Then
                        If (.Key > 0) Then
                            Set goSelected_ExtraLayer = graphicObject
                        End If
                    End If
                End With
            Next
        End If
    Next
    
'    For Each graphicObject In Esprit.Document.GraphicsCollection
'        With graphicObject
'            If (.Layer.Name = "STL" And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType = espSTL_Model)) Then
'                If (.Key > 0) Then
'                    Set goSelected = graphicObject
'                End If
'            End If
'        End With
'    Next
    
'    For Each graphicObject In Esprit.Document.GraphicsCollection
'        With graphicObject
'            If (.Layer.Name = "SPECIAL" And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType = espSTL_Model)) Then
'                If (.Key > 0) Then
'                    Set goSelected_ExtraLayer = graphicObject
'                End If
'            End If
'        End With
'    Next
    
    'Call fSelected.Add(goSelected, espFreeFormPartSurfaceItem)
    'Call fSelected.Remove()
    
    For Each fff In Application.Document.FreeFormFeatures
        If (fff.Name = "SPECIAL [SPECIAL]") Then
            Call ChangeFreeFormPartElement(fff, goSelected_ExtraLayer)
        Else
            Call ChangeFreeFormPartElement(fff, goSelected)
        End If
    Next

End Sub
Public Sub ChangeFreeFormPartElement(fSelected As Esprit.FreeFormFeature, goSelected As Esprit.graphicObject)
    Dim i As Integer
    Dim idxTobeDeleted As Integer
    Dim fffItem As EspritFeatures.ComFreeFormFeatureElement
    
    'For check goSelected is already set as an Element in the FFF
    Dim bAlreadySet As Boolean
    bAlreadySet = False
    
    ReDim ItemList(1 To fSelected.Count) As String
    'loop through items in FFF, compile a list of bodies (solids, surfaces) that make up that solid
    For i = 1 To fSelected.Count
        On Error Resume Next
        Set fffItem = fSelected.Item(i)
        Dim go As graphicObject, sb As EspritSolids.SolidBody, sf As EspritSolids.SolidFace
        Set go = Nothing
        If fffItem.Type = espFreeFormPartSurfaceItem Then
            If IsGraphicObj(fffItem.Object) Then
                Set go = fffItem.Object
            Else
                On Error Resume Next
                Set sb = fffItem.Object
                Set sf = fffItem.Object
                On Error GoTo 0
                If Not sb Is Nothing Then
                    Set go = GetGraphicObjectFromSolidBody(sb)
                End If
                If Not sf Is Nothing Then
                    Set go = GetGraphicObjectFromSolidFace(sf)
                End If
            End If
        
            'get graphic object solid id and add it to the list (if it doesnt exist already)
            If Not go Is Nothing Then
                'GoId = GO.TypeName & GO.Key & "(" & GO.Layer.Name & ")"
                '15 대체모델
                If go.Layer.Name = "대체모델" Then
                    idxTobeDeleted = i
                End If
                If go.Key = goSelected.Key Then
                    bAlreadySet = True
                End If
            End If
            On Error GoTo 0
        End If
    Next
    
    If bAlreadySet Then
        Exit Sub
    Else
        Call fSelected.Add(goSelected, espFreeFormPartSurfaceItem)
        Call fSelected.Remove(idxTobeDeleted)
        Call RebuildOperation(fSelected)
    End If
    
'    GetFreeformSolidIds = GetSimpleStringFromItemListArray(ItemList)
End Sub



Public Sub RebuildFreeformWithCheckElements()
'strCheckType
'all, 0, 90, 180, 270

    'Rebuild Operations
    Dim fff As Esprit.FreeFormFeature
    Dim fSelected As Esprit.FreeFormFeature
    Dim graphicObject As Esprit.graphicObject
    Dim goSelected As Esprit.graphicObject
    
    Dim str() As String
    Dim strParsefffName As String
    Dim strCheckLayer As String
    Dim x As Integer
    
    Call Application.OutputWindow.Clear
    For Each fff In Application.Document.FreeFormFeatures
        Call ClearFreeFormCheckElement(fff)
        strParsefffName = Mid(fff.Name, InStr(1, fff.Name, "[") + 1, InStr(1, fff.Name, "]") - InStr(1, fff.Name, "[") - 1)
        For x = LBound(Split(strParsefffName, "+")) To UBound(Split(strParsefffName, "+"))
            strCheckLayer = Split(strParsefffName, "+")(x)
            Call Application.OutputWindow.Text("FreeFormName: " & fff.Name & " / Parsing(" & CStr(x) & "): " & strCheckLayer & vbCrLf)
            For Each graphicObject In Esprit.Document.GraphicsCollection
                With graphicObject
                    'If (InStr(1, Replace(.Layer.Name, "", ""), strCheckLayer) <> 0 And (.GraphicObjectType = espSolidModel)) Then
                    If ((Replace(.Layer.Name, "", "") = strCheckLayer) And (.GraphicObjectType = espSolidModel)) Then
                        If (.Key > 0) Then
                            Set goSelected = graphicObject
                            Call ChangeFreeFormCheckElement(fff, goSelected)
                            Call Application.OutputWindow.Text("FreeFormName: " & fff.Name & " / CheckLayer: " & goSelected.Layer.Name & " / SolidNo: " & CStr(goSelected.Key) & vbCrLf)
                        End If
                    End If
                End With
            Next
        Next x
    Next
    
    'Call fSelected.Add(goSelected, espFreeFormPartSurfaceItem)
    'Call fSelected.Remove()
    
    

End Sub

Sub ClearFreeFormCheckElement(fSelected As Esprit.FreeFormFeature)
    Dim i As Integer
    Dim idxTobeDeleted As Integer
    Dim fffItem As EspritFeatures.ComFreeFormFeatureElement
    
    'For check goSelected is already set as an Element in the FFF
    Dim bAlreadySet As Boolean
    bAlreadySet = False
    Dim nRemoveItemIdxList() As Integer
    Dim x As Integer
    x = 0
    ReDim nRemoveItemIdxList(0 To fSelected.Count - 1)
    'loop through items in FFF, compile a list of bodies (solids, surfaces) that make up that solid
    For i = 1 To fSelected.Count
        On Error Resume Next
        Set fffItem = fSelected.Item(i)
        Dim go As graphicObject, sb As EspritSolids.SolidBody, sf As EspritSolids.SolidFace
        Set go = Nothing
        'Check Item
        If fffItem.Type = espFreeFormCheckSurfaceItem Then
            If IsGraphicObj(fffItem.Object) Then
                Set go = fffItem.Object
            Else
                On Error Resume Next
                Set sb = fffItem.Object
                Set sf = fffItem.Object
                On Error GoTo 0
                If Not sb Is Nothing Then
                    Set go = GetGraphicObjectFromSolidBody(sb)
                End If
                If Not sf Is Nothing Then
                    Set go = GetGraphicObjectFromSolidFace(sf)
                End If
            End If
        
            'get graphic object solid id and add it to the list (if it doesnt exist already)
            If Not go Is Nothing Then
                'GoId = GO.TypeName & GO.Key & "(" & GO.Layer.Name & ")"
                Call fSelected.Remove(i)
                i = 0
            End If
            On Error GoTo 0
        End If
    Next
    
    
End Sub
Private Sub ChangeFreeFormCheckElement(fSelected As Esprit.FreeFormFeature, goSelected As Esprit.graphicObject)
    Dim i As Integer
    Dim idxTobeDeleted As Integer
    Dim fffItem As EspritFeatures.ComFreeFormFeatureElement
    
    'For check goSelected is already set as an Element in the FFF
    Dim bAlreadySet As Boolean
    bAlreadySet = False
    
    ReDim ItemList(1 To fSelected.Count) As String
    'loop through items in FFF, compile a list of bodies (solids, surfaces) that make up that solid
    For i = 1 To fSelected.Count
        On Error Resume Next
        Set fffItem = fSelected.Item(i)
        Dim go As graphicObject, sb As EspritSolids.SolidBody, sf As EspritSolids.SolidFace
        Set go = Nothing
        'Check Item
        If fffItem.Type = espFreeFormCheckSurfaceItem Then
            If IsGraphicObj(fffItem.Object) Then
                Set go = fffItem.Object
            Else
                On Error Resume Next
                Set sb = fffItem.Object
                Set sf = fffItem.Object
                On Error GoTo 0
                If Not sb Is Nothing Then
                    Set go = GetGraphicObjectFromSolidBody(sb)
                End If
                If Not sf Is Nothing Then
                    Set go = GetGraphicObjectFromSolidFace(sf)
                End If
            End If
        
            'get graphic object solid id and add it to the list (if it doesnt exist already)
            If Not go Is Nothing Then
                'GoId = GO.TypeName & GO.Key & "(" & GO.Layer.Name & ")"
                '15 대체모델
                If go.Layer.Name = "대체모델" Then
                    idxTobeDeleted = i
                End If
                If go.Key = goSelected.Key Then
                    bAlreadySet = True
                End If
            End If
            On Error GoTo 0
        End If
    Next
    
    If bAlreadySet Then
        Exit Sub
    Else
        Call fSelected.Add(goSelected, espFreeFormCheckSurfaceItem)
        'Call fSelected.Remove(idxTobeDeleted)
'modified by Ian Pak at 2018.04.27 - R5
'        Call RebuildOperation(fSelected)
    End If
    
'    GetFreeformSolidIds = GetSimpleStringFromItemListArray(ItemList)
End Sub

Function CreateFreeFormFeaturesFromSTL(ByVal paramSTL As Esprit.STL_Model, Optional enableUndo As Boolean = True) As Esprit.FeatureSet
    Dim fff As Esprit.FreeFormFeature
    Dim fs As Esprit.FeatureSet ' folder on Features tab UI
    If enableUndo Then Document.OpenUndoTransaction
    
    Set fs = Document.FeatureSets.Add()
    fs.Name = "FeatureSet for STL " & paramSTL.Key
    Set fff = Document.FreeFormFeatures.Add
    fff.Name = "FreeFormFeature for STL " & paramSTL.Key
    Call fff.Add(paramSTL, espFreeFormPartSurfaceItem)
    Call fs.Add(fff)
    
    If enableUndo Then Call Document.CloseUndoTransaction(True)
    Set CreateFreeFormFeaturesFromSTL = fs

End Function

Function CreateFreeFormFeaturesFromLobes(ByVal sl As Esprit.Solid, Optional enableUndo As Boolean = True) As Esprit.FeatureSet
    Dim sb As EspritSolids.SolidBody
    Set sb = sl.SolidBody
    Dim sf As EspritSolids.SolidFace
    Dim fff As Esprit.FreeFormFeature
    Dim fs As Esprit.FeatureSet ' folder on Features tab UI
    If enableUndo Then Document.OpenUndoTransaction
    Set fs = Document.FeatureSets.Add()
    fs.Name = "Lobe Set for Solid " & sl.Key
    For Each sf In sb.SolidFaces
    If sf.SolidSurface.SurfaceType = geoSurfaceSwept Then
    Set fff = Document.FreeFormFeatures.Add
    fff.Name = "Lobe for Face " & sf.Identity
    Call fff.Add(sf, espFreeFormPartSurfaceItem)
    Call fs.Add(fff)
    End If
    Next
    If enableUndo Then Call Document.CloseUndoTransaction(True)
    Set CreateFreeFormFeaturesFromLobes = fs

End Function
Private Sub RemovePartElement()

End Sub

Private Sub RebuildOperation(fSelected As Esprit.FreeFormFeature)

'Application.Document.Operations.BatchMode = True

'Build Operations Here

'Application.Document.Operations.BatchMode = False

'Rebuild Operations
Dim Op As Esprit.Operation

'For Each fff In Application.Document.FreeFormFeatures
'    If fff.Name = "3 프리폼" Then
        For Each Op In Application.Document.Operations
             
              If Not (Op.Feature Is Nothing) Then
                  If Op.Feature.Name = fSelected.Name Then
                 'If op.Name = "8. 0 DEG CROSS BALL ENDMILL R0.75" Or op.Name = "8. 0 DEG -1 CROSS BALL ENDMILL R0.75" Then
                    DoEvents
                    Op.NeedsReexecute = True
                    DoEvents
                    'op.Grouped = False
                    'DoEvents
                    Op.Rebuild
                    DoEvents
                End If
            End If
        Next
'    End If
'Next

'release resource
Set Op = Nothing

End Sub
Sub SetBoundaryOperationAll()
    Call Application.OutputWindow.Clear
    Call Application.OutputWindow.Text("[SetBoundaryOperationAll] Begin" & vbCrLf)
    
    Call SetBoundaryOperation("0DEG")
    Call SetBoundaryOperation("90DEG")
    Call SetBoundaryOperation("180DEG")
    Call SetBoundaryOperation("270DEG")
    
    Call Application.OutputWindow.Text("[SetBoundaryOperationAll] End" & vbCrLf)


End Sub

Sub SetBoundaryOperation(pstrOpDEG As String)
'fSelected As Esprit.FreeFormFeature
'Application.Document.Operations.BatchMode = True

'Build Operations Here

'Application.Document.Operations.BatchMode = False


'    Dim FileName As String
'    FileName = "2_FRONT_TURNING.prc"
''    Dim feature As Esprit.FeatureChain
''    Set feature = Document.FeatureChains.Item(1)
'    Dim M_TechnologyUtility As EspritTechnology.TechnologyUtility
'    Set M_TechnologyUtility = Document.TechnologyUtility
'    Dim Tech() As EspritTechnology.Technology
'    Tech = M_TechnologyUtility.OpenProcess(EspritUserFolder & FileName)
'
'
'    Dim Op As Esprit.Operation
'
'    Set Op = Document.Operations.Add(Tech(0), Fc)
'    Op.Name = "2-1. FRONT TURNING"



    Dim strOpDEG, strLayerName, strLayerMarginName As String
    'strOpDEG = "0DEG"
    strOpDEG = pstrOpDEG
    strLayerName = strOpDEG & " CROSS BALL ENDMILL"
    strLayerMarginName = strOpDEG & " 마진"
    Dim Op As Esprit.Operation
    Dim strBoundaryPF As String
    strBoundaryPF = ""

    Dim M_TechnologyUtility As EspritTechnology.TechnologyUtility
    Set M_TechnologyUtility = Document.TechnologyUtility
    Dim techTLMPP As EspritTechnology.TechLatheMoldParallelPlanes
    Dim fff As Esprit.FreeFormFeature
    Dim Fc As Esprit.FeatureChain

'    Call Application.OutputWindow.Clear
    Call Application.OutputWindow.Text("[" & strOpDEG & "] Begin" & vbCrLf)
    
    For Each fff In Application.Document.FreeFormFeatures
        If (InStr(fff.Name, "[" & strOpDEG) > 0) Or (InStr(fff.Name, "+" & strOpDEG) > 0) Then
            For Each Op In Application.Document.Operations
                If Not (Op.Feature Is Nothing) Then
                    'Call Application.OutputWindow.Text(Op.Name & vbCrLf)
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '1) TechID withOUT "-1"
                    If (InStr(Op.Name, " " & strOpDEG) > 0 And InStr(Op.Name, " " & strOpDEG & "-1") <= 0) Then
                        strBoundaryPF = ""
                        For Each Fc In Document.FeatureChains
                            'If (Fc.Layer.Name = "STL" Or ly.Name = "기본값") Then
'                            If (Fc.Layer.Name = strLayerName) Then
'                                If strBoundaryPF = "" Then
'                                    strBoundaryPF = CStr(Fc.GraphicObjectType) & "," & CStr(Fc.Key)
'                                Else
'                                    strBoundaryPF = strBoundaryPF + "|" & CStr(Fc.GraphicObjectType) & "," & CStr(Fc.Key)
'                                End If
'                            End If
                            If (Fc.Layer.Name = strLayerMarginName) Then
                                If strBoundaryPF = "" Then
                                    strBoundaryPF = CStr(Fc.GraphicObjectType) & "," & CStr(Fc.Key)
                                Else
                                    strBoundaryPF = strBoundaryPF + "|" & CStr(Fc.GraphicObjectType) & "," & CStr(Fc.Key)
                                End If
                            End If
                        
                        Next
                        
                        Set techTLMPP = Op.Technology
                        'Call Application.OutputWindow.Text("[OperationName / ToolID]" & Op.Name & " / " & techTLMPP.ToolID & vbCrLf)
                        Call Application.OutputWindow.Text("[OperationName]" & Op.Name & vbCrLf)
                        Call Application.OutputWindow.Text("[Current BoundaryProfiles]" & techTLMPP.BoundaryProfiles & vbCrLf)
                        techTLMPP.BoundaryProfiles = strBoundaryPF
                        Call Application.OutputWindow.Text("[Updated BoundaryProfiles]" & techTLMPP.BoundaryProfiles & vbCrLf)
                    End If
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '2) TechID with "-1"
                    If (InStr(Op.Name, " " & strOpDEG & "-1") > 0) Then
                        strBoundaryPF = ""
                        For Each Fc In Document.FeatureChains
                            'If (Fc.Layer.Name = "STL" Or ly.Name = "기본값") Then
                            If (Fc.Layer.Name = strLayerName) Then
                                If strBoundaryPF = "" Then
                                    strBoundaryPF = CStr(Fc.GraphicObjectType) & "," & CStr(Fc.Key)
                                Else
                                    strBoundaryPF = strBoundaryPF + "|" & CStr(Fc.GraphicObjectType) & "," & CStr(Fc.Key)
                                End If
                            End If
                            If (Fc.Layer.Name = strLayerMarginName) Then
                                If strBoundaryPF = "" Then
                                    strBoundaryPF = CStr(Fc.GraphicObjectType) & "," & CStr(Fc.Key)
                                Else
                                    strBoundaryPF = strBoundaryPF + "|" & CStr(Fc.GraphicObjectType) & "," & CStr(Fc.Key)
                                End If
                            End If
                        
                        Next
                        
                        Set techTLMPP = Op.Technology
                        'Call Application.OutputWindow.Text("[OperationName / ToolID]" & Op.Name & " / " & techTLMPP.ToolID & vbCrLf)
                        Call Application.OutputWindow.Text("[OperationName]" & Op.Name & vbCrLf)
                        Call Application.OutputWindow.Text("[Current BoundaryProfiles]" & techTLMPP.BoundaryProfiles & vbCrLf)
                        techTLMPP.BoundaryProfiles = strBoundaryPF
                        Call Application.OutputWindow.Text("[Updated BoundaryProfiles]" & techTLMPP.BoundaryProfiles & vbCrLf)
                    End If
                
                End If
            Next
        End If
    Next

    Call Application.OutputWindow.Text("[" & strOpDEG & "] END" & vbCrLf)

    Set M_TechnologyUtility = Nothing
    Set techTLMPP = Nothing
    Set fff = Nothing
    Set Fc = Nothing

End Sub


Sub SetTextOperation(pstrOpDEG As String)

    Dim strOpDEG, strLayerName, strLayerMarginName As String
    'strOpDEG = "0DEG"
    strOpDEG = pstrOpDEG
    strLayerName = "TEXT"
    '??? strLayerMarginName
    strLayerMarginName = strOpDEG & " 마진"
    
    Dim Op As Esprit.Operation
    Dim strBoundaryPF As String
    strBoundaryPF = ""

    Dim M_TechnologyUtility As EspritTechnology.TechnologyUtility
    Set M_TechnologyUtility = Document.TechnologyUtility
    Dim techTLMPP As EspritTechnology.TechLatheMill3DProject
    Dim fff As Esprit.FreeFormFeature
    Dim Fc As Esprit.FeatureChain

'    Call Application.OutputWindow.Clear
    Call Application.OutputWindow.Text("[" & strOpDEG & "] Begin" & vbCrLf)
    
    For Each fff In Application.Document.FreeFormFeatures
        If (InStr(fff.Name, "[" & strOpDEG) > 0) Or (InStr(fff.Name, "+" & strOpDEG) > 0) Then
            For Each Op In Application.Document.Operations
                If Not (Op.Feature Is Nothing) Then
                    'Call Application.OutputWindow.Text(Op.Name & vbCrLf)
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '1) TechID withOUT "-1"
                    If (InStr(Op.Name, " " & strOpDEG) > 0 And InStr(Op.Name, " " & strOpDEG & "-1") <= 0) Then
                        strBoundaryPF = ""
                        For Each Fc In Document.FeatureChains
                            If (Fc.Layer.Name = strLayerMarginName) Then
                                If strBoundaryPF = "" Then
                                    strBoundaryPF = CStr(Fc.GraphicObjectType) & "," & CStr(Fc.Key)
                                Else
                                    strBoundaryPF = strBoundaryPF + "|" & CStr(Fc.GraphicObjectType) & "," & CStr(Fc.Key)
                                End If
                            End If
                        
                        Next
                        
                        Set techTLMPP = Op.Technology
                        'Call Application.OutputWindow.Text("[OperationName / ToolID]" & Op.Name & " / " & techTLMPP.ToolID & vbCrLf)
                        Call Application.OutputWindow.Text("[OperationName]" & Op.Name & vbCrLf)
                        Call Application.OutputWindow.Text("[Current BoundaryProfiles]" & techTLMPP.BoundaryProfiles & vbCrLf)
                        techTLMPP.BoundaryProfiles = strBoundaryPF
                        Call Application.OutputWindow.Text("[Updated BoundaryProfiles]" & techTLMPP.BoundaryProfiles & vbCrLf)
                    End If
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '2) TechID with "-1"
                    If (InStr(Op.Name, " " & strOpDEG & "-1") > 0) Then
                        strBoundaryPF = ""
                        For Each Fc In Document.FeatureChains
                            'If (Fc.Layer.Name = "STL" Or ly.Name = "기본값") Then
                            If (Fc.Layer.Name = strLayerName) Then
                                If strBoundaryPF = "" Then
                                    strBoundaryPF = CStr(Fc.GraphicObjectType) & "," & CStr(Fc.Key)
                                Else
                                    strBoundaryPF = strBoundaryPF + "|" & CStr(Fc.GraphicObjectType) & "," & CStr(Fc.Key)
                                End If
                            End If
                            If (Fc.Layer.Name = strLayerMarginName) Then
                                If strBoundaryPF = "" Then
                                    strBoundaryPF = CStr(Fc.GraphicObjectType) & "," & CStr(Fc.Key)
                                Else
                                    strBoundaryPF = strBoundaryPF + "|" & CStr(Fc.GraphicObjectType) & "," & CStr(Fc.Key)
                                End If
                            End If
                        
                        Next
                        
                        Set techTLMPP = Op.Technology
                        'Call Application.OutputWindow.Text("[OperationName / ToolID]" & Op.Name & " / " & techTLMPP.ToolID & vbCrLf)
                        Call Application.OutputWindow.Text("[OperationName]" & Op.Name & vbCrLf)
                        Call Application.OutputWindow.Text("[Current BoundaryProfiles]" & techTLMPP.BoundaryProfiles & vbCrLf)
                        techTLMPP.BoundaryProfiles = strBoundaryPF
                        Call Application.OutputWindow.Text("[Updated BoundaryProfiles]" & techTLMPP.BoundaryProfiles & vbCrLf)
                    End If
                
                End If
            Next
        End If
    Next

    Call Application.OutputWindow.Text("[" & strOpDEG & "] END" & vbCrLf)

    Set M_TechnologyUtility = Nothing
    Set techTLMPP = Nothing
    Set fff = Nothing
    Set Fc = Nothing

End Sub



Public Sub FreeFormCheckElementSolidRefresh(strOpDEGLayerName As String)
'strOpDEGLayerName = "0DEG 경계소재"
Dim fff As Esprit.FreeFormFeature
Dim graphicObject As Esprit.graphicObject
Dim goSelected As Esprit.graphicObject

'1. Find (fff) FreeFormFeature with "xxxDEG 경계소재"
Set fff = getFreeFormFeature(strOpDEGLayerName)
'2. Clear SolidModels of PartElement on the layer "xxxDEG 경계소재" in the fff
Call FreeFormCheckElementDelete(fff, strOpDEGLayerName)
'3. Add SolidModels of PartElement on the layer "xxxDEG 경계소재" in the fff
For Each graphicObject In Esprit.Document.GraphicsCollection
    With graphicObject
        If ((Replace(.Layer.Name, "", "") = strOpDEGLayerName) And (.GraphicObjectType = espSolidModel)) Then
            If (.Key > 0) Then
                Set goSelected = graphicObject
                Call FreeFormCheckElementAdd(fff, goSelected)
            End If
        End If
    End With
Next

Set fff = Nothing
Set goSelected = Nothing

'For Each fff In Application.Document.FreeFormFeatures
'    If (InStr(fff.Name, "[" & strOpDEGLayerName) > 0) Or (InStr(fff.Name, "+" & strOpDEGLayerName) > 0) Then
'        Call GetFreeformSolidIds(fff)
'        For Each graphicObject In Esprit.Document.GraphicsCollection
'            With graphicObject
'                If ((Replace(.Layer.Name, "", "") = strOpDEGLayerName) And (.GraphicObjectType = espSolidModel)) Then
'                    If (.Key > 0) Then
'                        Set goSelected = graphicObject
'                        Call ChangeFreeFormCheckElement(fff, goSelected)
'                        Call Application.OutputWindow.Text("FreeFormName: " & fff.Name & " / CheckLayer: " & goSelected.Layer.Name & " / SolidNo: " & CStr(goSelected.Key) & vbCrLf)
'                    End If
'                End If
'            End With
'        Next
'    End If
'Next

End Sub

Public Function getFreeFormFeature(pStrFindWordInName As String) As FreeFormFeature
    Dim rtnFFF As Esprit.FreeFormFeature
    Dim fff As Esprit.FreeFormFeature
    'Dim nCount As int
    'nCount = 0

    For Each fff In Application.Document.FreeFormFeatures
        If (InStr(fff.Name, "[" & pStrFindWordInName) > 0) Or (InStr(fff.Name, "+" & pStrFindWordInName) > 0) Then
            Set rtnFFF = fff
            'nCount = nCount + 1
        End If
    Next
    
    Set getFreeFormFeature = rtnFFF
End Function

Public Sub FreeFormCheckElementDelete(fSelected As Esprit.FreeFormFeature, pLayerName As String)
    Dim i As Integer
    Dim idxTobeDeleted As Integer
    Dim fffItem As EspritFeatures.ComFreeFormFeatureElement
    
    Dim go As graphicObject, sb As EspritSolids.SolidBody, sf As EspritSolids.SolidFace
    Set go = Nothing
    
    'loop through items in FFF, compile a list of bodies (solids, surfaces) that make up that solid
    For i = 1 To fSelected.Count
        On Error Resume Next
        Set fffItem = fSelected.Item(i)
        If fffItem.Type = espFreeFormCheckSurfaceItem Then
            If IsGraphicObj(fffItem.Object) Then
                Set go = fffItem.Object
            Else
                On Error Resume Next
                Set sb = fffItem.Object
                Set sf = fffItem.Object
                On Error GoTo 0
                If Not sb Is Nothing Then
                    Set go = GetGraphicObjectFromSolidBody(sb)
                End If
                If Not sf Is Nothing Then
                    Set go = GetGraphicObjectFromSolidFace(sf)
                End If
            End If
        
            'Remove graphic objects in the Layer from fff.
            If Not go Is Nothing Then
                If go.Layer.Name = pLayerName Then
                    idxTobeDeleted = i
                    Call fSelected.Remove(idxTobeDeleted)
                    Call Application.OutputWindow.Text("FreeFormName: " & fSelected.Name & " / CheckLayer: " & go.Layer.Name & " / Deleted SolidModel No: " & CStr(go.Key) & vbCrLf)
                End If
            End If
            On Error GoTo 0
        End If
    Next
    
    Set go = Nothing
    Set fffItem = Nothing
End Sub

Public Sub FreeFormCheckElementAdd(fSelected As Esprit.FreeFormFeature, goSelected As Esprit.graphicObject)
    Call fSelected.Add(goSelected, espFreeFormCheckSurfaceItem)
    Call Application.OutputWindow.Text("FreeFormName: " & fSelected.Name & " / CheckLayer: " & goSelected.Layer.Name & " / SolidNo: " & CStr(goSelected.Key) & vbCrLf)
End Sub

