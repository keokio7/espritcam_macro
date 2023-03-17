VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSTLRotate 
   Caption         =   "STL Rotate"
   ClientHeight    =   2550
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

Private Sub cmdbtnDirectionCheck_Click()
    RunDirectionCheck
End Sub

Public Sub RunDirectionCheck(Optional ByVal pnAutoTry As Integer = 0)
    If CheckSTLDirectionLine Then
        Call MsgBox("Right direction.", vbOKOnly, "Direction Check")
    Else
        If pnAutoTry = 0 Then
            If MsgBox("Wrong direction. Try to rotate and test again.", vbYesNoCancel, "Direction Check") = vbYes Then
                Call TurnSTLByDirectionCheckInfo
                RunDirectionCheck
            End If
        Else
            For i = 1 To pnAutoTry
                Call TurnSTLByDirectionCheckInfo
                If CheckSTLDirectionLine Then
                    Call MsgBox("Right direction.", vbOKOnly, "Direction Check")
                    Exit For
                End If
            Next i
        End If
    End If
    
    Dim ly As Esprit.Layer
    For Each ly In Document.Layers
        If (ly.Name = "STL" Or ly.Name = "기본값" Or ly.Name = "BACK TURNING" Or ly.Name = "CUT-OFF" Or ly.Name = "CUF-OFF" Or (InStr(1, ly.Name, "방향체크") <> 0)) Then
            ly.Visible = True
        Else
            ly.Visible = False
        End If
    Next
    Document.Refresh
    
End Sub


Public Sub TurnSTLByDirectionCheckInfo()
'HEX[H]
'KEY WAY[K]
'OCTA[O]
'SQUARE[S]
'TORX[X]
'TRIANGLE[T]
    Dim strTemp As String
    strTemp = ""
    
    strTemp = GetConnection("Angle")
    If IsNumeric(strTemp) Then
        Debug.Print "Try to Turn: X, " + strTemp
        Call TurnSTL("X," + strTemp)
    Else
        strTemp = GetConnection("Style")
        Select Case strTemp
        Case "H" 'HEX
            Debug.Print "Try to Turn: HEX Default X,30"
            Call TurnSTL("X,30")
        Case "K" 'KEY WAY[K]
            Debug.Print "Try to Turn: KEY WAY Default X,30"
            Call TurnSTL("X,30")
        Case "O" 'OCTA[O]
            Debug.Print "Try to Turn: OCTA Default X,22.5"
            Call TurnSTL("X,22.5")
        Case "S" 'SQUARE[S]
            Debug.Print "Try to Turn: SQUARE Default X,45"
            Call TurnSTL("X,45")
        Case "X" 'TORX[X]
            Debug.Print "Try to Turn: TORX Default X,30"
            Call TurnSTL("X,30")
        Case "T" 'TRIANGLE[T]
            Debug.Print "Try to Turn: TRIANGLE Default X,60"
            Call TurnSTL("X,60")
        Case Else    ' Other values.
            Debug.Print "Not in H,K,O,S,X,T. " + strTemp
        End Select
    End If
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
