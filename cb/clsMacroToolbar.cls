VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMacroToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'clsMacroToolbar
' ------------------------------------------------------------------------
' Copyright (C) 1998-2006 DP Technology Corp.
'
' This sample code file is provided for illustration purposes. You have a
' royalty-free right to use, modify, reproduce and distribute the sample
' code files (and/or any modified versions) in any way you find useful,
' provided that you agree that DP Technology has no warranty obligations
' or liability whatsoever for any of the sample code files, the result
' from your use of them or for any modifications you may make. This notice
' must be included on any reproduced or distributed file or modified file.
' ------------------------------------------------------------------------

Option Explicit

Const NUM_MACROS = 7 ' number of macros, change as desired
Const TOOLBAR_NAME = "CAM Automation" ' change as desired

Dim WithEvents AddIn As EspritCommands.AddIn
Attribute AddIn.VB_VarHelpID = -1

Dim MyCookie As Long
Dim MyCommand() As Long
Dim CommandActive() As Boolean

Private Sub AddIn_OnCommand(ByVal Cookie As Long, ByVal UserId As Long)
    If Cookie = MyCookie Then
        CommandActive(UserId) = True
        Select Case UserId
        ' provide as many case statements as necessary ( = NUM_MACROS )
        Case 1
            ' insert code to call macro #  1 here
            Call ClickBtn1
        Case 2
            ' insert code to call macro #  2 here
            Call ClickBtn2
        Case 3
            ' insert code to call macro #  3 here
            Call ClickBtn3
        Case 4
            ' insert code to call macro #  4 here
            Call ClickBtn4
        Case 5
            ' insert code to call macro #  5 here
            Call ClickBtn5
        Case 6
            ' insert code to call macro #  T here
            Call ClickBtnT
        Case 7
            ' insert code to call macro #  R here
            Call ClickBtnR
            
        End Select
        CommandActive(UserId) = False
    End If
End Sub

'Private Sub AddIn_OnCommandUI(ByVal Cookie As Long, ByVal UserId As Long, Enabled As Boolean, Checked As Boolean)
'    ' this keeps the button depressed while macro is executing
'    ' this event procedure can be removed if desired
'    If Cookie = MyCookie Then
'        Checked = CommandActive(UserId)
'    End If
'End Sub

Private Sub Class_Initialize()
    Set AddIn = Application.AddIn
    MyCookie = AddIn.GetCookie
    '
    ' resize the arrays based on the number of macros and initialize the command values for each
    '
    ReDim MyCommand(1 To NUM_MACROS) As Long
    ReDim CommandActive(1 To NUM_MACROS) As Boolean
    Dim i As Long
    For i = 1 To NUM_MACROS
        MyCommand(i) = AddIn.AddCommand(MyCookie, i, "Command" & i)
    Next
    '
    ' remove then recreate the toolbar and add the command buttons
    '
    On Error Resume Next ' go on to line after error in case it does not exist
    Call Application.ToolBars.Remove(TOOLBAR_NAME)
    On Error GoTo 0
    Dim TB As Esprit.ToolBar
    Set TB = Application.ToolBars.Add(TOOLBAR_NAME)
    With TB
        For i = 1 To NUM_MACROS
            Call .Add(espToolBarControlButton, "Command " & str(i), MyCommand(i))
            .Item(i).Enabled = True
        Next
        '
        ' set the remaining toolbar button parameters
        '
        .Visible = True
        '
        ' provide as many ScreenTip assignments as necessary ( = NUM_MACROS )
        '
        .Item(1).ScreenTip = "#1 to set a STL and Connection in a proper location. #1-2 to make the Solid Turn Toolpaths [FRONT TURNING]"
        .Item(2).ScreenTip = "#2 to generate toolpaths for [ROUGH ENDMILL R6.0]. Please make sure the STL properly located."
        .Item(3).ScreenTip = "#3 to create 경계소재2 - border material 2nd"
        .Item(4).ScreenTip = "#4 to rebuild freeforms"
        .Item(5).ScreenTip = "#5 to create Margin"
        .Item(6).ScreenTip = "#T to engrave program number"
        .Item(7).ScreenTip = "#R to reorder operations by name"
        '
        ' provide as many HelpString assignments as necessary ( = NUM_MACROS )
        '
        .Item(1).HelpString = "#1"
        .Item(2).HelpString = "#2"
        .Item(3).HelpString = "#3"
        .Item(4).HelpString = "#4"
        .Item(5).HelpString = "#5"
        .Item(6).HelpString = "#T"
        .Item(7).HelpString = "#R"
        
        '
        ' if you wish to have an image on the button, then on the lines below modify the
        ' base BmpPath and the path names given for the .bmp files to something appropriate
        ' the example shown is for images named 16-1.bmp, 32-1.bmp, 16-2.bmp, 32-2.bmp, etc.
        ' with those images kept in a Bitmaps subfolder under the default Templates folder.
        '
        Dim BmpPath As String
        BmpPath = Application.Configuration.GetFileDirectory(espFileTypeTemplate) & "\Bitmaps\"
        Call Application.OutputWindow.Clear
        Call Application.OutputWindow.Text("[Button file path] " & BmpPath & vbCrLf)
        
        On Error Resume Next ' go on to line after error in case bitmaps do not exist
        'For i = 1 To NUM_MACROS
        '    Call .Item(i).SetBitmap(BmpPath & "16-" & i & ".bmp", BmpPath & "32-" & i & ".bmp")
        'Next
        
        Call .Item(1).SetBitmap(BmpPath & "16-V2-1.bmp", BmpPath & "32-V2-1.bmp")
        Call .Item(2).SetBitmap(BmpPath & "16-V2-2.bmp", BmpPath & "32-V2-2.bmp")
        Call .Item(3).SetBitmap(BmpPath & "16-V2-3.bmp", BmpPath & "32-V2-3.bmp")
        Call .Item(4).SetBitmap(BmpPath & "16-V2-4.bmp", BmpPath & "32-V2-4.bmp")
        Call .Item(5).SetBitmap(BmpPath & "16-V2-5.bmp", BmpPath & "32-V2-5.bmp")
        Call .Item(6).SetBitmap(BmpPath & "16-V2-T.bmp", BmpPath & "32-V2-T.bmp")
        Call .Item(7).SetBitmap(BmpPath & "16-V2-R.bmp", BmpPath & "32-V2-R.bmp")
        
        .Dock (espToolBarPositionTop)
    End With
    
    
End Sub

Private Sub Class_Terminate()
    ' remove the toolbar if it exists
    On Error Resume Next ' go on to line after error in case it does not exist
    Call Application.ToolBars.Remove(TOOLBAR_NAME)
    Set AddIn = Nothing
End Sub


