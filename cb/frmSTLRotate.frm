VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSTLRotate 
   Caption         =   "STL Rotate"
   ClientHeight    =   2055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2175
   OleObjectBlob   =   "frmSTLRotate.frx":0000
End
Attribute VB_Name = "frmSTLRotate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub chkFace_Click()
    
    Dim strWorkPlane As String
    strWorkPlane = ""
    
    If chkFace Then
        Call SaveSetting("frmSTLRotate", "SaveWorkPlane", "SaveWorkPlane", Document.ActivePlane.Name)
        Document.ActivePlane = Document.Planes("FACE")
    Else
        strWorkPlane = GetSetting("frmSTLRotate", "SaveWorkPlane", "SaveWorkPlane", 0)
        If strWorkPlane = "" Then strWorkPlane = "0DEG"
        
        Document.ActivePlane = Document.Planes(strWorkPlane)
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

Private Sub cmdBtn2_Click()
    Call ClickBtn1_2
End Sub

Private Sub cmdRotate045_Click()
    Call TurnSTL("X,45")
End Sub

Private Sub cmdRotate060_Click()
    Call TurnSTL("X,60")
End Sub

Private Sub cmdRotate090_Click()
    Call TurnSTL("X,90")
End Sub

Private Sub cmdRotate120_Click()
    Call TurnSTL("X,120")
End Sub
