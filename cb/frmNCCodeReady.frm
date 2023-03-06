VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNCCodeReady 
   Caption         =   "Is it ready to NC Code?"
   ClientHeight    =   525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2760
   OleObjectBlob   =   "frmNCCodeReady.frx":0000
End
Attribute VB_Name = "frmNCCodeReady"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
   Me.Left = GetSetting("Userform Positioning", "Position-Left-" + Me.Name, "Left", 0)
   Me.Top = GetSetting("Userform Positioning", "Position-Top-" + Me.Name, "Top", 0)
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   Call SaveSetting("Userform Positioning", "Position-Left-" + Me.Name, "Left", Me.Left)
   Call SaveSetting("Userform Positioning", "Position-Top-" + Me.Name, "Top", Me.Top)
End Sub

Private Sub cmdCancel_Click()
    frmNCCodeReady.Hide
    Unload frmNCCodeReady
End Sub

Private Sub cmdConfirm_Click()
    frmNCCodeReady.Hide
    Unload frmNCCodeReady
    Call showAdvancedNCCode
End Sub


