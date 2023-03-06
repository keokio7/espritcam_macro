Attribute VB_Name = "autoToolbar"
' ------------------------------------------------------------------------
' Copyright (C) 1998-2007 DP Technology Corp.
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

Dim MacroToolbarDemo As clsMacroToolbar

' after importing this module check to see if you have other auto run procedures
' in your global project; having duplicate auto run procedures can cause conflicts
'
Sub EspritStartup()
    ' if you have another EspritStartup procedure in your global project then
    ' include a call to MacroToolbarInitialize in it and delete this procedure
    MacroToolbarInitialize
End Sub

Sub AutoOpen()
    ' if you have another AutoOpen procedure in your global project then
    ' include a call to MacroToolbarInitialize in it and delete this procedure
    MacroToolbarInitialize
End Sub

Sub EspritShutdown()
    ' if you have another EspritShutdown procedure in your global project then
    ' include a call to MacroToolbarCancel in it and delete this procedure
    MacroToolbarCancel
End Sub

Sub MacroToolbarInitialize()
    ' to create the toolbar if it was not already created
    If MacroToolbarDemo Is Nothing Then Set MacroToolbarDemo = New clsMacroToolbar
End Sub

Sub MacroToolbarCancel()
    ' to automatically remove the toolbar at shutdown
    Set MacroToolbarDemo = Nothing
End Sub

Sub ToolbarButtonDemo(MacroNumber As Long)
    ' this is just a place holder macro to give the buttons something to do
    ' other macros can go in this or other standard code modules
    ' call those macros from the AddIn_OnCommand event in clsMacroToolbar
    Call MsgBox("You pressed the button for macro # " & MacroNumber, vbOKOnly, "CAM Automation")
End Sub
