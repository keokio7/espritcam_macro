Attribute VB_Name = "Macro"
'Macro_Main
Option Explicit
 
Public Const DEFAULT_TOLERANCE = "0.01"
 
Public Const CSIDL_DESKTOP = &H0   ' Desktop (namespace root)
Public Const CSIDL_DESKTOPDIRECTORY = &H10 ' Desktop folder ([user] profile)
Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19 ' Desktop folder (All Users profile)
Public Const MAX_PATH = 260
Public Const NOERROR = 0

Public Type shiEMID
    cb As Long
    abID As Byte
End Type
Public Type ITEMIDLIST
    mkid As shiEMID
End Type

Public Const HWND_TOPMOST = -1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOREPOSITION = &H200
Public Const SWP_NOSIZE = &H1

Public Enum WindowState
   SW_HIDE = 0
   SW_SHOWNORMAL = 1
   SW_NORMAL = 1
   SW_SHOWMINIMIZED = 2
   SW_SHOWMAXIMIZED = 3
   SW_MAXIMIZE = 3
   SW_SHOWNOACTIVATE = 4
   SW_SHOW = 5
   SW_MINIMIZE = 6
   SW_SHOWMINNOACTIVE = 7
   SW_SHOWNA = 8
   SW_RESTORE = 9
   SW_SHOWDEFAULT = 10
   SW_FORCEMINIMIZE = 11
   SW_MAX = 11
End Enum

Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function GetSpecialfolder(CSIDL As Long) As String
    Dim IDL As ITEMIDLIST
    Dim sPath As String
    Dim iReturn As Long
    
    iReturn = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    
    If iReturn = NOERROR Then
        sPath = Space(512)
        iReturn = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        sPath = RTrim$(sPath)
        If Asc(Right(sPath, 1)) = 0 Then sPath = Left$(sPath, Len(sPath) - 1)
        GetSpecialfolder = sPath
        Exit Function
    End If
    GetSpecialfolder = ""
End Function
Sub ResetLayer(strLayerName As String)
    Dim ly As Esprit.Layer
    For Each ly In Document.Layers
        If (ly.Name = strLayerName) Then
            Call Document.Layers.Remove(strLayerName)
            Exit For
        End If
    Next
    
    Call Document.Layers.Add(strLayerName)

End Sub
Sub AddLayersForAutomation()
    ResetLayer ("STL")
    ResetLayer ("0DEG 경계소재")
    ResetLayer ("90DEG 경계소재")
    ResetLayer ("180DEG 경계소재")
    ResetLayer ("270DEG 경계소재")
    ResetLayer ("DummyOperation")
    ResetLayer ("0DEG 마진")
    ResetLayer ("90DEG 마진")
    ResetLayer ("180DEG 마진")
    ResetLayer ("270DEG 마진")

End Sub

 
Function GetWorkFolder() As String
    Dim strDesk As String
    strDesk = GetSpecialfolder(CSIDL_DESKTOP)
    GetWorkFolder = strDesk + "\작업\"
    'or
    'strDesk = GetSpecialFolder(CSIDL_DESKTOPDIRECTORY)
    'or
    'strDesk = GetSpecialFolder(CSIDL_COMMON_DESKTOPDIRECTORY)
End Function
Function GetWorkScanfileFolder() As String
    GetWorkScanfileFolder = GetWorkFolder + "스캔파일\"
End Function
Function GetWorkEspritfileFolder() As String
    GetWorkEspritfileFolder = GetWorkFolder + "작업저장\"
End Function
 
 ''need to set reference to Windows Script Host Object Model
Function CopyFileToBackup(strFileNameFrom As String, strFileNameTo As String, strPathFrom As String, strPathTo As String) As Boolean
    Dim objFSO As FileSystemObject
    Dim objFolder As Folder
    Dim objFile As File
    
    CopyFileToBackup = True
    On Error GoTo Err_handler
    
    Set objFSO = New FileSystemObject
    If strFileNameTo = "" Then strFileNameTo = strFileNameFrom
    
    'Set objFolder = objFSO.GetFolder("E:\Esprit\개발관련\Step1\STLFiles\A\")
    Set objFolder = objFSO.GetFolder(strPathFrom)
    Set objFile = objFSO.GetFile(strPathFrom + strFileNameFrom)
    
    objFile.Copy strPathTo, True
    MsgBox ("The STL file has been backed up successfully to " + strPathTo + strFileNameFrom)

    Set objFSO = Nothing
    Exit Function

Err_handler:
    MsgBox Err.Number & "-" & Err.Description
    Set objFSO = Nothing
    CopyFileToBackup = False

End Function
Function GetFilenameWithoutExtension(ByVal FileName)
  Dim Result, i
  Result = FileName
  i = InStrRev(FileName, ".")
  If (i > 0) Then
    Result = Mid(FileName, 1, i - 1)
  End If
  GetFilenameWithoutExtension = Result
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' BTN#1 [1] Select STL file & locate properly.
Public Sub ClickBtn1()
    
    On Error GoTo 0
    
    If MsgBox("[1]Select STL file & locate properly.", vbYesNo, "CAM Automation") = vbYes Then
        
        Dim strOriginalEsppritFileName As String
        strOriginalEsppritFileName = Document.FileName
        
        'FindOrientationShortestDimension
        'FindOrientationSmallestArea
        Dim strSTLFilePath As String
        'strSTLFilePath = ".\Avaneer_Tech_01_Williams_Hiossen reg_0.stl"
        
        Dim File As stcFileStruct
        '// fill values (not required)
        File.strDialogtitle = "Select file to open"
        File.strFilter = "STL files (*.stl)|*.stl|All files (*.*)|*.*" '// use same format as
        
        'Task1 stl 스캔 파일 open 시 자동으로 기본설정 파일(.esp)과 해당 stl 파일 불러오기
        'Set default STL file path
        If (Strings.Right(Document.Name, 4) = ".esp") Then
            File.strFileName = Strings.Left(Document.Name, Strings.Len(Document.Name) - 4) + ".stl"
        Else
            File.strFileName = Document.Name + ".stl"
        End If
    
        If Dir(GetWorkFolder, vbDirectory) = "" Then
            If MsgBox(GetWorkFolder + " is not found. Do you want to make the folder?", vbYesNo) = vbYes Then
                MkDir GetWorkFolder
            End If
            If Dir(GetWorkScanfileFolder, vbDirectory) = "" Then
                If MsgBox(GetWorkScanfileFolder + " is not found. Do you want to make the folder?", vbYesNo) = vbYes Then
                    MkDir GetWorkScanfileFolder
                End If
            End If
            If Dir(GetWorkEspritfileFolder, vbDirectory) = "" Then
                If MsgBox(GetWorkEspritfileFolder + " is not found. Do you want to make the folder?", vbYesNo) = vbYes Then
                    MkDir GetWorkEspritfileFolder
                End If
            End If
            MsgBox ("Workfolders has been generated. Please check folders and files and try it again.")
            Exit Sub
        End If
    
    
        Dim strFileExists As String
        strFileExists = Dir(GetWorkScanfileFolder + File.strFileName)
        
        If strFileExists = "" Then
        'The selected file doesn't exist
        'Get the file manually.
            ShowOpenDialog File
            If File.strFileName = "" Then
                Call MsgBox("Please select an STL file and try it again.")
                Exit Sub
            End If
            strSTLFilePath = File.strFileName
            Document.SaveAs (GetWorkEspritfileFolder + GetFilenameWithoutExtension(File.strFileTitle) + ".esp")
            
            If strOriginalEsppritFileName = Document.FileName Then
                Exit Sub
            End If
        Else
            strSTLFilePath = GetWorkScanfileFolder + File.strFileName
        End If
    
        'CommonDialog Control
        '// pass stcFileStruct
        '// get return values (passed back through type)
        'strSTLFilePath = ".\Core Dental Studio_16731_1_Hiossen ET Regular_pm.stl"
        
        
        'If Not CopyFileToBackup(File.strFileTitle, "", Replace(File.strFileName, File.strFileTitle, ""), "E:\Esprit\개발관련\Step1\STLFiles\B\") Then
        'Task1 stl 스캔 파일 open 시 자동으로 기본설정 파일(.esp)과 해당 stl 파일 불러오기
        'esp file already copied and open with external program.
        'Document.SaveAs (GetWorkEspritfileFolder + GetFilenameWithoutExtension(File.strFileTitle) + ".esp")
        'If Then
        '    MsgBox ("Back-up The STL file is failed. Please check it and try it again.")
        '    'Exit Sub
        'End If
        
        
        'If strOriginalEsppritFileName = Document.FileName Then
        '    Exit Sub
        'End If

        
        Dim strTurning As String
        strTurning = ""
        'Call GetSTL(strSTLFilePath)
        Call CheckSTL(strSTLFilePath)
        Document.Refresh
        
        Dim cCondition As Boolean
        cCondition = True
        Do While cCondition
            Select Case MsgBox("Is the STL properly located?", vbYesNoCancel)
                Case vbYes
                    If CheckSTLInTheCircle() Then
                        Call SelectSTL_Model
                        cCondition = False
                    End If
                Case vbNo
                    strTurning = InputBox("Enter X or Y and degree like (X,90).", "CAM Automation - To transform the STL", "X,0")
                    Call TurnSTL(strTurning)
                    cCondition = True
                Case vbCancel
                    cCondition = False
                Case Else
                    cCondition = False
            End Select
        Loop
    Else
        Exit Sub
    End If


    Dim ly As Esprit.Layer
    For Each ly In Document.Layers
        If (ly.Name = "STL" Or ly.Name = "기본값") Then
            ly.Visible = True
        End If
    Next

    For Each ly In Document.Layers
        If (ly.Name = "경계소재-1") Then
            ly.Visible = False
        End If
    Next
    Document.Refresh

    Load frmSTLRotate
    Call frmSTLRotate.RunDirectionCheck(1)
    frmSTLRotate.Show (vbModeless)

    Exit Sub
    
Err_handler:
    MsgBox Err.Number & "-" & Err.Description
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' BTN#1-2 [1-2] Generate toolpaths for [FRONT TURNING]. Please make sure the STL properly located."
Public Sub ClickBtn1_2()

    Load frmSTLRotate
    frmSTLRotate.Hide
    Unload frmSTLRotate
   

    If MsgBox("[2] Generate toolpaths for [FRONT TURNING]. Please make sure the STL properly located.", vbYesNo, "CAM Automation") = vbYes Then
        Call Step1_2
       Call Step2_4
        Call Step2_6
        
        Dim ly As Esprit.Layer
        For Each ly In Document.Layers
            If (ly.Name = "STL" Or ly.Name = "±?º?°ª") Then
                ly.Visible = True
            End If
        Next
        
    Else
        Exit Sub
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' BTN#2 [2] Generate toolpaths for [ROUGH ENDMILL R6.0]. Please make sure the STL properly located."
Public Sub ClickBtn2()
    If MsgBox("[2] Generate toolpaths for [ROUGH ENDMILL R6.0]. Please make sure the STL properly located.", vbYesNo, "CAM Automation") = vbYes Then
        
        
        Dim nRtn_0Deg As Integer
        Dim nRtn_120Deg As Integer
        Dim nRtn_240Deg As Integer
        nRtn_0Deg = 0
        nRtn_120Deg = 0
        nRtn_240Deg = 0
        
        If MsgBox("It is processing to 0DEG.", vbYesNo) = vbYes Then
            nRtn_0Deg = generateSolidmilTurn("0DEG", "ROUGH ENDMILL R6.0", "1")
        End If
        If MsgBox("It is processing to 120DEG.", vbYesNo) = vbYes Then
            nRtn_120Deg = generateSolidmilTurn("120DEG", "ROUGH ENDMILL R6.0", "2")
        End If
        If MsgBox("It is processing to 240DEG.", vbYesNo) = vbYes Then
            nRtn_240Deg = generateSolidmilTurn("240DEG", "ROUGH ENDMILL R6.0", "3")
        End If
    
        Dim ly As Esprit.Layer
        For Each ly In Document.Layers
            If (ly.Name = "STL" Or ly.Name = "기본값") Then
                ly.Visible = True
            End If
        Next
    
        If (nRtn_0Deg = 1 And nRtn_120Deg = 1 And nRtn_240Deg = 1) Then
            If MsgBox("Reorder Operation and show checking Rough Endmill.", vbYesNo) = vbYes Then
                Call ReorderOperation
                Unload frmCreateBorderSolidObject
                Load frmCreateBorderSolidObject
                frmCreateBorderSolidObject.MultiPage1.Value = 0
                
                'Show 선반소재(MaskLatheStock)
                Call Document.Windows.ActiveWindow.SetMask(espViewMaskLatheStock, True)
                Document.Refresh
                frmCreateBorderSolidObject.Show (vbModeless)
                
            End If
        End If
    
    Else
        Exit Sub
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' BTN#3 [3] Create 경계소재2 - border material 2nd
Public Sub ClickBtn3()
    Unload frmCreateBorderSolidObject
    Load frmCreateBorderSolidObject
    frmCreateBorderSolidObject.MultiPage1.Value = 1
    frmCreateBorderSolidObject.Show (vbModeless)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' BTN#4 Rebuild Freeform
Public Sub ClickBtn4()
    If MsgBox("[4] Rebuild Freeform With Part & Check Elements.", vbYesNo, "CAM Automation") = vbYes Then
        Call SetBoundaryOperationAll
        Call RebuildFreeformWithCheckElements
        Call RebuildFreeformWithNewSTLAsAPartElement
    Else
        Exit Sub
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' BTN#5 [5] Create Margin
Public Sub ClickBtn5()
    Unload frmCreateMargin
    Load frmCreateMargin
    frmCreateMargin.Show (vbModeless)
End Sub


Public Sub ClickBtnR()
    If MsgBox("[R] Reorder operations.", vbYesNo, "CAM Automation") = vbYes Then
        Call ReorderOperation
        Unload frmNCCodeReady
        Load frmNCCodeReady
        frmNCCodeReady.Show (vbModeless)
    Else
        Exit Sub
    End If
End Sub

Public Sub ClickBtnT()
    Dim strPGMNumber As String
    strPGMNumber = GetProgramNumber()
    '1. 데이터 페이지에 프로그램번호 입력
    Document.LatheMachineSetup.ProgramNumber = strPGMNumber
    
    Dim hdTemp As Esprit.Head
    For Each hdTemp In Document.LatheMachineSetup.Heads
        hdTemp.ProgramNumber = strPGMNumber
    Next
    
    Document.MillMachineSetup.ProgramNumber = strPGMNumber
    
    '2. Engraving Program Number
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
    
    Unload frmPGMText
    Load frmPGMText
    frmPGMText.Show (vbModeless)
    
End Sub


Public Sub ClickBtn3_Err()
    If MsgBox("[3] Generate toolpaths for [ROUGH ENDMILL R6.0]. Please make sure the STL properly located.", vbYesNo, "CAM Automation") = vbYes Then
        
        
        Dim nRtn As Integer
        nRtn = 0
        If MsgBox("It is processing to 0DEG.", vbYesNo) = vbYes Then
            nRtn = generateSolidmilTurn("0DEG", "ROUGH ENDMILL R6.0", "1", 0.001)
        End If
        If MsgBox("It is processing to 120DEG.", vbYesNo) = vbYes Then
            Call generateSolidmilTurn("120DEG", "ROUGH ENDMILL R6.0", "2", 0.001)
        End If
        If MsgBox("It is processing to 240DEG.", vbYesNo) = vbYes Then
            Call generateSolidmilTurn("240DEG", "ROUGH ENDMILL R6.0", "3", 0.001)
        End If
    
        Dim ly As Esprit.Layer
        For Each ly In Document.Layers
            If (ly.Name = "STL" Or ly.Name = "기본값") Then
                ly.Visible = True
            End If
        Next
    
    Else
        Exit Sub
    End If
End Sub

Private Function CheckSTLInTheCircle() As Boolean
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
    
    If (CheckSTLInTheCircle And Not (cBound Is Nothing)) Then
        Call Document.GraphicsCollection.Remove(cBound.GraphicsCollectionIndex)
    End If
    
    Document.Refresh


End Function


Public Sub OnTop(hwnd As Long)
    '
    ' put hWnd always on top
    '
    Call SetWindowPos(hwnd, -1, 0, 0, 0, 0, &H2 Or &H1)

End Sub


Public Sub OffTop(hwnd As Long)

    Call SetWindowPos(hwnd, -2, 0, 0, 0, 0, &H2 Or &H1)

End Sub

Public Function GetProgramNumber() As String

    Dim strFileName As String
    Dim strCodes() As String
    Dim strCode As String
    Dim strBaseESPCode As String
    Dim strPGMCode As String

    strFileName = Document.FileName
    strCodes = Split(strFileName, "(")

    strCode = strCodes(UBound(strCodes))
    strCodes = Split(strCode, ")")
    strCode = strCodes(LBound(strCodes))

    strCodes = Split(strCode, ",")
    strBaseESPCode = strCodes(LBound(strCodes))
    strPGMCode = strCodes(UBound(strCodes))
    
    'PGMCode Check Logic recommended.

    GetProgramNumber = strPGMCode

End Function


Public Function CheckSTLDirectionLine() As Boolean
'Is The STL in the Circle?

Call GetPartProfileSTL(0.01)

'Function IntersectCircleAndArcsSegments(ByRef cBound As Esprit.Circle, _
'            ByRef go2 As Esprit.graphicObject) As Esprit.Point

    Dim sTestSegment As Esprit.Segment
    Dim go2 As Esprit.graphicObject
    Dim pntIntersect As Esprit.Point
    Dim layerObject As Esprit.Layer
    Dim graphicObject As Esprit.graphicObject
        
    Dim seg As Esprit.Segment
    For Each seg In Esprit.Document.Segments
        With seg
        'If (.Layer.Name = "방향체크") Then
        If InStr(1, .Layer.Name, "방향체크") <> 0 Then
            If (.Key > 0) Then
                Set sTestSegment = seg
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
                If (pntIntersect Is Nothing) Then
                    Set pntIntersect = IntersectSegmentAndArcsSegments(sTestSegment, go2)
                    Call Document.GraphicsCollection.Remove(go2.GraphicsCollectionIndex)
                Else
                    Call Document.GraphicsCollection.Remove(go2.GraphicsCollectionIndex)
                End If
                Document.Refresh
                
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
    
    'Dim bInTheCircle As Boolean
    CheckSTLDirectionLine = Not (pntIntersect Is Nothing)
    
    'If (CheckSTLDirectionLine And Not (sTestSegment Is Nothing)) Then
    '    Call Document.GraphicsCollection.Remove(sTestSegment.GraphicsCollectionIndex)
    'End If
    
    Document.Refresh


End Function

Public Function GetConnection(pstrReturnClass As String) As String
    
    Dim strReturn As String
    Dim strConnectionType As String
    strReturn = ""
    strConnectionType = ""
    
    If HasDirectionCheckSegmentLayer Then
        Select Case pstrReturnClass
        Case "Style"
            strReturn = GetConnectionStyle
        Case "Angle"
            strConnectionType = GetConnectionStyle
            If InStr(1, strConnectionType, "Error_") Then
                strReturn = strConnectionType
            Else
                strReturn = GetTurningAngle
            End If
        Case Else
            strReturn = "Error_NotSupportedReturnClass"
        End Select
    Else
        strReturn = "Error_MissingDirectionCheckSegmentLayer"
    End If
    
    GetConnection = strReturn

End Function
Private Function HasDirectionCheckSegmentLayer() As Boolean
    
    Dim lyrObject As Esprit.Layer
    HasDirectionCheckSegmentLayer = False
    
    For Each lyrObject In Esprit.Document.Layers
        With lyrObject
            If InStr(1, .Name, "방향체크") <> 0 Then
                HasDirectionCheckSegmentLayer = True
                Exit For
            End If
        End With
    Next

End Function

Private Function GetConnectionStyle() As String
    
    Dim lyrObject As Esprit.Layer
    Dim strConnectionType As String
    strConnectionType = ""
    
    For Each lyrObject In Esprit.Document.Layers
        With lyrObject
        If InStr(1, .Name, "방향체크") <> 0 Then
            strConnectionType = Mid(Replace(.Name, "방향체크", ""), 1, 1)
            Select Case strConnectionType
            Case "H"
            Case "K"
            Case "O"
            Case "S"
            Case "X"
            Case "T"
                strConnectionType = strConnectionType
            Case Else
                strConnectionType = "Error_NotInRegeisteredConnectionStyleCode"
            End Select

            Exit For
        End If
        End With
    Next
    
    GetConnectionStyle = strConnectionType

End Function

Private Function GetTurningAngle() As String
    
    Dim lyrObject As Esprit.Layer
    Dim strTurningAngle As String
    strTurningAngle = ""
    
    For Each lyrObject In Esprit.Document.Layers
        With lyrObject
        If InStr(1, .Name, "방향체크") <> 0 Then
            strTurningAngle = Replace(.Name, "방향체크" + GetConnectionStyle, "")
            If Not IsNumeric(strTurningAngle) Then
                strTurningAngle = "Error_IsNotNumber"
            End If
            Exit For
        End If
        End With
    Next
    
    GetTurningAngle = strTurningAngle

End Function

Public Function IsAlpha(s) As Boolean
    IsAlpha = Len(s) And Not s Like "*[!a-zA-Z]*"
End Function

