VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'cGroup (Class Module)
Sub ScanGroup()

    Application.OutputWindow.Clear

    Application.OutputWindow.Visible = True

    Dim G As Esprit.graphicObject

    Dim i As Long

    Call Application.OutputWindow.Text("Group contains " & Document.Group.Count & " GraphicObject items" & vbCrLf)

    For i = 1 To Document.Group.Count

        Set G = Document.Group.Item(i)

        If G.GraphicObjectType <> espUnknown Then

            Call Application.OutputWindow.Text("Item " & i & " is " & G.TypeName & G.Key & vbCrLf)

        End If

    Next

End Sub



Sub FeatureChainGroupedCirclesAndArcs()

    Dim PI As Double

    PI = 4 * Atn(1)

    Dim A As Esprit.Arc

    Dim C As Esprit.Circle

    Dim Fc As Esprit.FeatureChain

    Dim Angle As Double

    Angle = InputBox("Enter Starting Angle", "Starting Angle", 90)

    Angle = Angle * PI / 180 'convert to Radians

    '

    ' remember current settings

    '

    Dim ActiveClr As Long

    ActiveClr = Document.ActiveColor

    Document.ActiveColor = vbBlue

    '

    ' loop through the Group

    '

    Dim G As Esprit.graphicObject

    Dim i As Long

    For i = 1 To Document.Group.Count

        Set G = Document.Group.Item(i)

        Select Case G.GraphicObjectType

        Case espArc

            Set A = G

            If (Abs(A.EndAngle - A.StartAngle) = 2 * PI) Then

                '

                ' create the feature chain

                '

                Set Fc = Document.FeatureChains.Add(A.Extremity(espExtremityStart))

                Call Fc.Add(A)

                Fc.Layer = A.Layer

            End If

        Case espCircle

            Set C = G

            '

            ' create the feature chain

            '

            Set Fc = Document.FeatureChains.Add(Document.GetPoint(C.x + C.Radius * Cos(Angle), C.y + C.Radius * Sin(Angle), C.Z))

            Fc.Layer = C.Layer

            '

            ' because you cannot add a circle to a FeatureChain, make temporary arc instead

            '

            Call Fc.Add(Document.GetArc(C.CenterPoint, C.Radius, Angle, 2 * PI + Angle))

        End Select

    Next

    '

    ' reset to original value

    '

    Document.ActiveColor = ActiveClr

    Document.Refresh

End Sub




Sub FeatureChainGroupedCirclesAndArcs2()

    Dim PI As Double

    PI = 4 * Atn(1)

    Dim A As Esprit.Arc

    Dim C As Esprit.Circle

    Dim Fc As Esprit.FeatureChain

    Dim Angle As Double

    Angle = InputBox("Enter Starting Angle", "Starting Angle", 90)

    Angle = Angle * PI / 180 'convert to Radians

    '

    ' remember current settings

    '

    Dim ActiveClr As Long

    ActiveClr = Document.ActiveColor

    Document.ActiveColor = vbBlue

    '

    ' loop through all the circles

    '

    For Each C In Document.Circles

        If C.Grouped Then

            '

            ' create the feature chain

            '

            Set Fc = Document.FeatureChains.Add(Document.GetPoint(C.x + C.Radius * Cos(Angle), C.y + C.Radius * Sin(Angle), C.Z))

            Fc.Layer = C.Layer

            '

            ' because you cannot add a circle to a FeatureChain, make temporary arc instead

            '

            Call Fc.Add(Document.GetArc(C.CenterPoint, C.Radius, Angle, 2 * PI + Angle))

        End If

    Next

    '

    ' loop through all the arcs

    '

    For Each A In Document.Arcs

        If A.Grouped Then

            If (Abs(A.EndAngle - A.StartAngle) = 2 * PI) Then

                '

                ' create the feature chain

                '

                Set Fc = Document.FeatureChains.Add(A.Extremity(espExtremityStart))

                Call Fc.Add(A)

                Fc.Layer = A.Layer

            End If

        End If

    Next

    '

    ' reset to original value

    '

    Document.ActiveColor = ActiveClr

    Call FitAllWindows

    Document.Refresh

End Sub

 

Sub FitAllWindows()

    Dim W As Esprit.Window

    For Each W In Document.Windows

        W.Fit

        '

        ' note that Fit just sizes the window to fit the graphics

        ' it does not redraw the graphics, so we still need a Refresh

        '

        W.Refresh

    Next

End Sub





