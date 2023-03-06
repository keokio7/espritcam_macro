Attribute VB_Name = "test"
'test
'Update History Log
    'Modified by Ian Pak at 2021.02.10
    'If cannot define M1_Segment, retry with round 3


Private Const dBottom = 0
Private Const dTop = 1

Public Enum udfPlaneDegree
    udfPlaneDegree0 = 0
    udfPlaneDegree10 = 10
    udfPlaneDegree20 = 20
    udfPlaneDegree30 = 30
    udfPlaneDegree40 = 40
    udfPlaneDegree50 = 50
    udfPlaneDegree60 = 60
    udfPlaneDegree70 = 70
    udfPlaneDegree80 = 80
    udfPlaneDegree90 = 90
    udfPlaneDegree100 = 100
    udfPlaneDegree110 = 110
    udfPlaneDegree120 = 120
    udfPlaneDegree130 = 130
    udfPlaneDegree140 = 140
    udfPlaneDegree150 = 150
    udfPlaneDegree160 = 160
    udfPlaneDegree170 = 170
    udfPlaneDegree180 = 180
    udfPlaneDegree190 = 190
    udfPlaneDegree200 = 200
    udfPlaneDegree210 = 210
    udfPlaneDegree220 = 220
    udfPlaneDegree230 = 230
    udfPlaneDegree240 = 240
    udfPlaneDegree250 = 250
    udfPlaneDegree260 = 260
    udfPlaneDegree270 = 270
    udfPlaneDegree280 = 280
    udfPlaneDegree290 = 290
    udfPlaneDegree300 = 300
    udfPlaneDegree310 = 310
    udfPlaneDegree320 = 320
    udfPlaneDegree330 = 330
    udfPlaneDegree340 = 340
    udfPlaneDegree350 = 350
End Enum


Private Sub ScanAnnotationsAdvanced()

    Application.OutputWindow.Clear

    Application.OutputWindow.Visible = True

    Dim An As Esprit.Annotation

    Dim i As Long

    For i = 1 To Document.Annotations.Count

        Set An = Document.Annotations.Item(i)

        '

        ' type cast the annonation to a specific object based on the type

        '

        Select Case An.AnnotationType

        Case espAnnotationDimension

            Call Application.OutputWindow.Text("Item " & i & " is Dimension " & An.Key & vbCrLf)

            Dim d As Esprit.Dimension

            Set d = An

            Call Application.OutputWindow.Text("Dimension " & d.Key & " contains " & d.AnnotationSegments.Count & " segment(s) and " & d.Notes.Count & " note(s)." & vbCrLf)

        Case espAnnotationHatch

            Call Application.OutputWindow.Text("Item " & i & " is Hatch " & An.Key & vbCrLf)

            Dim H As Esprit.Hatch

            Set H = An

            Call Application.OutputWindow.Text("Hatch " & H.Key & " contains " & H.AnnotationSegments.Count & " segments." & vbCrLf)

        Case espAnnotationLeader

            Call Application.OutputWindow.Text("Item " & i & " is Leader " & An.Key & vbCrLf)

            Dim Ld As Esprit.Leader

            Set Ld = An

            Call Application.OutputWindow.Text("Leader " & Ld.Key & " contains " & Ld.LeaderPoints.Count & " leader point(s)." & vbCrLf)

        Case espAnnotationNote

            Call Application.OutputWindow.Text("Item " & i & " is Note " & An.Key & vbCrLf)

            Dim n As Esprit.Notes

            Set n = An

            Call Application.OutputWindow.Text("Note " & n.Key & " contains " & n.Count & " line(s) of text." & vbCrLf)

        End Select

    Next

End Sub

Function Step2_4() As Integer
'4) 3)에서 이동한 선분을 왼쪽 방향으로 2 만큼 연장했을 때
'  STL 개체와 만나는가?
'  - Y : 선분을 그대로 수평으로 STL 개체와 만날 때까지 연장
'  - N : 선분을 좌측아래 방향 45도로 STL 개체와 만날 때까지 연장On Error Resume Next

'Return
'1: Success
'-991: Cannot find a parallel segment.
'-999: Other Error

    'Declare the variables
    Dim App As Esprit.Application
    Dim doc As Esprit.Document
    Dim M1_Segment As Esprit.Segment
    Dim M1_SegmentEx As Esprit.Segment
    
    Dim M2_Segment As Esprit.Segment
    Dim M2_Arc As Esprit.Arc
    
    Dim pnt_Point As Esprit.Point
    Dim sg_ResultSegment As Esprit.Segment
    Dim strBaseLayerName As String
    strBaseLayerName = "FRONT TURNING"
    
    'Initialize the App and Doc Variables
    Set App = Application
    Set doc = App.Document

    Dim r(2), STLRightEndX  As Double
    Dim minX As Double
    minX = 999
    
    r(1) = GetCutOffXRightEnd
    r(2) = GetSTLXEnd(1)
    STLRightEndX = r(2)
    'Modify an existing Segment
    '자동으로 찾게
    For Each segmentObject In Esprit.Document.Segments
        With segmentObject
        If (.Layer.Name = strBaseLayerName) Then
            '수평인지 확인: (Round(.YStart, 5) = Round(.YEnd, 5))
            'STL Rigght End 이전에 있는지 확인
            If ((Round(.YStart, 5) = Round(.YEnd, 5)) And (.XStart < .XEnd And .XStart < minX) And (.XStart <= STLRightEndX And STLRightEndX < .XEnd)) Then
                If (.Key > 0) Then
                    minX = .XStart
                    .Grouped = True
                    .Layer.Visible = True
                    Set M1_Segment = segmentObject
                End If
            ElseIf ((Round(.YStart, 5) = Round(.YEnd, 5)) And (.XStart > .XEnd And .XEnd < minX) And (.XEnd <= STLRightEndX And STLRightEndX < .XStart)) Then
                If (.Key > 0) Then
                    minX = .XEnd
                    .Grouped = True
                    .Layer.Visible = True
                    Set M1_Segment = segmentObject
                End If
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
    
    
    'Modified by Ian Pak at 2021.02.10
    'If cannot define M1_Segment, retry with round 3
    If M1_Segment Is Nothing Then
        For Each segmentObject In Esprit.Document.Segments
            With segmentObject
            If (.Layer.Name = strBaseLayerName) Then
                '수평인지 확인: (Round(.YStart, 3) = Round(.YEnd,3))
                'STL Rigght End 이전에 있는지 확인
                If ((Round(.YStart, 3) = Round(.YEnd, 3)) And (.XStart < .XEnd And .XStart < minX) And (.XStart <= STLRightEndX And STLRightEndX < .XEnd)) Then
                    If (.Key > 0) Then
                        minX = .XStart
                        .Grouped = True
                        .Layer.Visible = True
                        Set M1_Segment = segmentObject
                    End If
                ElseIf ((Round(.YStart, 3) = Round(.YEnd, 3)) And (.XStart > .XEnd And .XEnd < minX) And (.XEnd <= STLRightEndX And STLRightEndX < .XStart)) Then
                    If (.Key > 0) Then
                        minX = .XEnd
                        .Grouped = True
                        .Layer.Visible = True
                        Set M1_Segment = segmentObject
                    End If
                End If
            ElseIf (.GraphicObjectType <> espUnknown) Then
                .Grouped = False
            End If
            End With
        Next
    End If
    
    'Modified by Ian Pak at 2021.09.10
    'If cannot define M1_Segment, retry with round 2
    If M1_Segment Is Nothing Then
        For Each segmentObject In Esprit.Document.Segments
            With segmentObject
            If (.Layer.Name = strBaseLayerName) Then
                '수평인지 확인: (Round(.YStart, 2) = Round(.YEnd,2))
                'STL Right End 이전에 있는지 확인
                If ((Round(.YStart, 2) = Round(.YEnd, 2)) And (.XStart < .XEnd And .XStart < minX) And (.XStart <= STLRightEndX And STLRightEndX < .XEnd)) Then
                    If (.Key > 0) Then
                        minX = .XStart
                        .Grouped = True
                        .Layer.Visible = True
                        Set M1_Segment = segmentObject
                    End If
                ElseIf ((Round(.YStart, 2) = Round(.YEnd, 2)) And (.XStart > .XEnd And .XEnd < minX) And (.XEnd <= STLRightEndX And STLRightEndX < .XStart)) Then
                    If (.Key > 0) Then
                        minX = .XEnd
                        .Grouped = True
                        .Layer.Visible = True
                        Set M1_Segment = segmentObject
                    End If
                End If
            ElseIf (.GraphicObjectType <> espUnknown) Then
                .Grouped = False
            End If
            End With
        Next
    End If
    
    'Modified by Ian Pak at 2021.09.27
    'If cannot define M1_Segment, retry with round 2
    If M1_Segment Is Nothing Then
        For Each segmentObject In Esprit.Document.Segments
            With segmentObject
            If (.Layer.Name = strBaseLayerName) Then
                '수평인지 확인: (Abs(.YStart - .YEnd) < 0.01)
                'STL Right End 이전에 있는지 확인
                If ((Abs(.YStart - .YEnd) < 0.01) And (.XStart < .XEnd And .XStart < minX) And (.XStart <= STLRightEndX And STLRightEndX < .XEnd)) Then
                    If (.Key > 0) Then
                        minX = .XStart
                        .Grouped = True
                        .Layer.Visible = True
                        Set M1_Segment = segmentObject
                    End If
                ElseIf ((Abs(.YStart - .YEnd) < 0.01) And (.XStart > .XEnd And .XEnd < minX) And (.XEnd <= STLRightEndX And STLRightEndX < .XStart)) Then
                    If (.Key > 0) Then
                        minX = .XEnd
                        .Grouped = True
                        .Layer.Visible = True
                        Set M1_Segment = segmentObject
                    End If
                End If
            ElseIf (.GraphicObjectType <> espUnknown) Then
                .Grouped = False
            End If
            End With
        Next
    End If
    
    
    'Modified by Ian Pak at 2021.09.13
    'If cannot define M1_Segment, retry with round 2
    If M1_Segment Is Nothing Then
        'Call MsgBox("Cannot find a parallel segment.", vbCritical, "Error in Front Turning")
        Step2_4 = -991
        Exit Function
    End If
    
'    Set M1_Segment = doc.Segments("122")
    If (M1_Segment.XStart > M1_Segment.XEnd) Then
        M1_Segment.Reverse
    End If
    Set M1_SegmentEx = M1_Segment
    M1_SegmentEx.XStart = M1_SegmentEx.XStart - 2

    'Debug
    M1_SegmentEx.Grouped = True
    'M1_Segment.XStart = M1_Segment.XStart - 2
    
'    Set M2_Segment = doc.Segments("126")
    
    For Each segmentObject In Esprit.Document.Segments
        Set M2_Segment = segmentObject
        With segmentObject
        If (.Layer.Name = strBaseLayerName And M2_Segment.Key <> M1_SegmentEx.Key) Then
            'Debug
            M2_Segment.Grouped = True
            
            'If ((maxValue(.YStart, .YEnd) > 0) And (minValue(.XStart, .XEnd) < M1_Segment.XStart) And (maxValue(.XStart, .XEnd) > M1_SegmentEx.XStart)) Then
            If ((maxValue(.YStart, .YEnd) > 0) And (minValue(.XStart, .XEnd) < M1_Segment.XEnd) And (maxValue(.XStart, .XEnd) > M1_SegmentEx.XStart)) Then
                If (.Key > 0) Then
                    'Set pntIntersect = IntersectArcsSegments(M1_Segment, M2_Segment)
                    Set pntIntersect = IntersectArcsSegments(M1_SegmentEx, M2_Segment)
                    If Not (pntIntersect Is Nothing) Then
                        'For Debug
                        pntIntersect.Grouped = True
                        .Layer.Visible = True
                        If (M2_Segment.XStart > M2_Segment.XEnd) Then
                            M2_Segment.Reverse
                        End If
                        M2_Segment.XEnd = pntIntersect.x
                        M2_Segment.YEnd = pntIntersect.y
                        'M2_Segment.ZEnd = pntIntersect.Z
                        Exit For
                    End If
                End If
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
'*** Later it should be trimed for Arcs / Modified Ian at 2022.03.02.
    If pntIntersect Is Nothing Then
        For Each arcObject In Esprit.Document.Arcs
            Set M2_Arc = arcObject
            With arcObject
            If (.Layer.Name = strBaseLayerName) Then
                If ((maxValue(.Extremity(espExtremityStart).y, .Extremity(espExtremityEnd).y) > 0) _
                    And (minValue(.Extremity(espExtremityStart).x, .Extremity(espExtremityEnd).x) < M1_Segment.XEnd) _
                    And (maxValue(.Extremity(espExtremityStart).x, .Extremity(espExtremityEnd).x) > M1_SegmentEx.XStart)) Then
                    If (.Key > 0) Then
                        Set pntIntersect = IntersectArcsSegments(M1_SegmentEx, M2_Arc)
                        'For Debug
                        If Not (pntIntersect Is Nothing) Then
                            'For Debug
                            pntIntersect.Grouped = True
                            .Layer.Visible = True
                            If (M2_Arc.Extremity(espExtremityStart).x > M2_Arc.Extremity(espExtremityEnd).x) Then
                                M2_Arc.Reverse
                            End If
                            M2_Arc.Extremity(espExtremityEnd).x = pntIntersect.x
                            M2_Arc.Extremity(espExtremityEnd).y = pntIntersect.y
                            'M2_Arc.Extremity(espExtremityEnd).Z = pntIntersect.Z
                            Exit For
                        End If
                    End If
                End If
            ElseIf (.GraphicObjectType <> espUnknown) Then
                .Grouped = False
            End If
            End With
        Next
    End If
    
    If Not (pntIntersect Is Nothing) Then
        M1_Segment.XStart = pntIntersect.x
        If ((M1_Segment.YStart - pntIntersect.y) < 0.01) Then
            M1_Segment.YStart = pntIntersect.y
        End If
        ''''''''''''''''''''''''''''''''''''''''''
        'Trim below the extended segment
        Dim tmpSegment As Esprit.Segment
        Dim tmpArc As Esprit.Arc
        For Each goObject In Esprit.Document.GraphicsCollection
        With goObject
        If (.Layer.Name = strBaseLayerName) Then
            If (.GraphicObjectType = espSegment) Then
                Set tmpSegment = goObject
                tmpSegment.Grouped = True
                With tmpSegment
                    If ((minValue(.XStart, .XEnd) > pntIntersect.x) And (maxValue(.YStart, .YEnd) < pntIntersect.y)) Then
                        Call Document.GraphicsCollection.Remove(.GraphicsCollectionIndex)
                        Document.Refresh
                    End If
                End With
                tmpSegment.Grouped = False
            ElseIf (.GraphicObjectType = espArc) Then
                Set tmpArc = goObject
                tmpArc.Grouped = True
                
                With tmpArc
                    If ((minValue(.Extremity(espExtremityStart).x, .Extremity(espExtremityEnd).x) > pntIntersect.x) _
                        And (maxValue(.Extremity(espExtremityStart).y, .Extremity(espExtremityEnd).y) < pntIntersect.y)) Then
                        Call Document.GraphicsCollection.Remove(.GraphicsCollectionIndex)
                        Document.Refresh
                    End If
                End With
                tmpArc.Grouped = False
            
            Else
'                .Grouped = False
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next

        
    Else
        'Set sg_ResultSegment = Document.GetSegment(Document.GetPoint(M1_Segment.XStart - Sqr(2), M1_Segment.YStart - Sqr(2), M1_Segment.ZStart), Document.GetPoint(M1_Segment.XStart, M1_Segment.YStart, M1_Segment.ZStart))
        Set sg_ResultSegment = Document.Segments.Add(Document.GetPoint(M1_Segment.XStart - Sqr(2), M1_Segment.YStart - Sqr(2), M1_Segment.ZStart), Document.GetPoint(M1_Segment.XStart, M1_Segment.YStart, M1_Segment.ZStart))
Document.Refresh

    '    Set sg_ResultSegment = doc.Segments("122")
        If (sg_ResultSegment.XStart > sg_ResultSegment.XEnd) Then
            sg_ResultSegment.Reverse
        End If

        For Each segmentObject In Esprit.Document.Segments
            Set M2_Segment = segmentObject
            With segmentObject
            If (.Layer.Name = strBaseLayerName) Then
                If ((maxValue(.YStart, .YEnd) > 0) And (minValue(.XStart, .XEnd) < sg_ResultSegment.XEnd) And (maxValue(.XStart, .XEnd) > sg_ResultSegment.XStart)) Then
                    If (.Key > 0) Then
                        Set pntIntersect = IntersectArcsSegments(sg_ResultSegment, M2_Segment)
                        If Not (pntIntersect Is Nothing) Then
                            'For Debug
                            'pntIntersect.Grouped = True
                            '.Layer.Visible = True
                            If (M2_Segment.XStart > M2_Segment.XEnd) Then
                                M2_Segment.Reverse
                            End If
                            M2_Segment.XEnd = pntIntersect.x
                            M2_Segment.YEnd = pntIntersect.y
                            'M2_Segment.ZEnd = pntIntersect.Z
                            Exit For
                        End If
                    End If
                End If
            ElseIf (.GraphicObjectType <> espUnknown) Then
                .Grouped = False
            End If
            End With
        Next
        'If pntIntersect Is Nothing Then
            For Each arcObject In Esprit.Document.Arcs
                Set M2_Arc = arcObject
                With arcObject
                If (.Layer.Name = strBaseLayerName) Then
                    If ((maxValue(.Extremity(espExtremityStart).y, .Extremity(espExtremityEnd).y) > 0) _
                        And (minValue(.Extremity(espExtremityStart).x, .Extremity(espExtremityEnd).x) < sg_ResultSegment.XEnd) _
                        And (maxValue(.Extremity(espExtremityStart).x, .Extremity(espExtremityEnd).x) > sg_ResultSegment.XStart)) Then
                        If (.Key > 0) Then
                            Set pntIntersect = IntersectArcsSegments(sg_ResultSegment, M2_Arc)
                            'For Debug
                            If Not (pntIntersect Is Nothing) Then
                                'For Debug
                                'pntIntersect.Grouped = True
                                '.Layer.Visible = True
                                If (M2_Arc.Extremity(espExtremityStart).x > M2_Arc.Extremity(espExtremityEnd).x) Then
                                    M2_Arc.Reverse
                                End If
                                M2_Arc.Extremity(espExtremityEnd).x = pntIntersect.x
                                M2_Arc.Extremity(espExtremityEnd).y = pntIntersect.y
                                'M2_Arc.Extremity(espExtremityEnd).Z = pntIntersect.Z
                                Exit For
                            End If
                        End If
                    End If
                ElseIf (.GraphicObjectType <> espUnknown) Then
                    .Grouped = False
                End If
                End With
            Next
        'End If
        
        If Not (pntIntersect Is Nothing) Then
            sg_ResultSegment.XStart = pntIntersect.x
            sg_ResultSegment.YStart = pntIntersect.y
            
            ''''''''''''''''''''''''''''''''''''''''''
            'Trim below the extended segment
            For Each goObject In Esprit.Document.GraphicsCollection
            With goObject
            If (.Layer.Name = strBaseLayerName) Then
                If (.GraphicObjectType = espSegment) Then
                    Set tmpSegment = goObject
                    tmpSegment.Grouped = True
                    With tmpSegment
'Modified by Ian Pak at 2018.04.03
                        'If ((minValue(.XStart, .XEnd) >= sg_ResultSegment.XStart) And (minValue(.XStart, .XEnd) <= sg_ResultSegment.XEnd) And (maxValue(.YStart, .YEnd) < sg_ResultSegment.YEnd)) Then
                        If ((minValue(.XStart, .XEnd) >= sg_ResultSegment.XStart) And (maxValue(.YStart, .YEnd) < sg_ResultSegment.YEnd)) Then
                            Call Document.GraphicsCollection.Remove(.GraphicsCollectionIndex)
                            Document.Refresh
                        End If
                    End With
                    tmpSegment.Grouped = False
                ElseIf (.GraphicObjectType = espArc) Then
                    Set tmpArc = goObject
                    tmpArc.Grouped = True
                    
                    With tmpArc
'Modified by Ian Pak at 2018.04.03
'                        If ((minValue(.Extremity(espExtremityStart).X, .Extremity(espExtremityEnd).X) >= sg_ResultSegment.XStart) _
'                            And (minValue(.Extremity(espExtremityStart).X, .Extremity(espExtremityEnd).X) <= sg_ResultSegment.XEnd) _
'                            And (maxValue(.Extremity(espExtremityStart).Y, .Extremity(espExtremityEnd).Y) < sg_ResultSegment.YEnd)) Then
                        If ((minValue(.Extremity(espExtremityStart).x, .Extremity(espExtremityEnd).x) >= sg_ResultSegment.XStart) _
                            And (maxValue(.Extremity(espExtremityStart).y, .Extremity(espExtremityEnd).y) < sg_ResultSegment.YEnd)) Then
                            
                            Call Document.GraphicsCollection.Remove(.GraphicsCollectionIndex)
                            Document.Refresh
                        End If
                    End With
                    tmpArc.Grouped = False
                
                Else
    '                .Grouped = False
                End If
            ElseIf (.GraphicObjectType <> espUnknown) Then
                .Grouped = False
            End If
            End With
            Next
        End If

    End If



Document.Refresh

'Trim 경계소재 부분
    For Each segmentObject In Esprit.Document.Segments
        Set M2_Segment = segmentObject
        With segmentObject
        If (.Layer.Name = strBaseLayerName) Then
            If ((maxValue(.YStart, .YEnd) > 0) And (minValue(.XStart, .XEnd) >= M1_Segment.XEnd)) Then
                If (.Key > 0) Then
                    Set pntIntersect = IntersectArcsSegments(M1_Segment, M2_Segment)
                    If Not (pntIntersect Is Nothing) Then
                        'For Debug
                        'pntIntersect.Grouped = True
                        '.Layer.Visible = True
                        If (M2_Segment.YStart < M2_Segment.YEnd) Then
                            M2_Segment.Reverse
                        End If
                        M2_Segment.XEnd = pntIntersect.x
                        M2_Segment.YEnd = pntIntersect.y
                        'M2_Segment.ZEnd = pntIntersect.Z
                        Exit For
                    End If
                End If
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
    'If pntIntersect Is Nothing Then
        For Each arcObject In Esprit.Document.Arcs
            Set M2_Arc = arcObject
            With arcObject
            If (.Layer.Name = strBaseLayerName) Then
                If ((maxValue(.Extremity(espExtremityStart).y, .Extremity(espExtremityEnd).y) > 0) _
                    And (minValue(.Extremity(espExtremityStart).x, .Extremity(espExtremityEnd).x) >= M1_Segment.XEnd)) Then
                    If (.Key > 0) Then
                        Set pntIntersect = IntersectArcsSegments(M1_Segment, M2_Arc)
                        'For Debug
                        If Not (pntIntersect Is Nothing) Then
                            'For Debug
                            'pntIntersect.Grouped = True
                            '.Layer.Visible = True
                            If (M2_Arc.Extremity(espExtremityStart).y < M2_Arc.Extremity(espExtremityEnd).y) Then
                                M2_Arc.Reverse
                            End If
                            M2_Arc.Extremity(espExtremityEnd).x = pntIntersect.x
                            M2_Arc.Extremity(espExtremityEnd).y = pntIntersect.y
                            'M2_Arc.Extremity(espExtremityEnd).Z = pntIntersect.Z
                            Exit For
                        End If
                    End If
                End If
            ElseIf (.GraphicObjectType <> espUnknown) Then
                .Grouped = False
            End If
            End With
        Next
    'End If



    Document.Refresh
    'Very Important - destroy all of the Objects
    Set M_Segment = Nothing
    Set doc = Nothing
    Set App = Nothing
    
    Step2_4 = 1 'Success Code
End Function



Function Step2_6() As Integer
'연결된 선을 [자동연결]로 연결 피처를 생성
On Error Resume Next


'Return
'1: Success
'-991: More than 2 Chain features are made.
'-999: Other Error

    Dim graphicObject As Esprit.graphicObject
    Dim nCount As Integer
    Dim lyTemp As Esprit.Layer
    
    nCount = Document.Group.Count
    nCount = Document.SelectionSets.Count
   
'2. Select ;
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("FrontTurning")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("FrontTurning")
    End With
    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
    End With
    
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Number = 1 And .Layer.Visible And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType = espArc Or .GraphicObjectType = espPoint Or .GraphicObjectType = espSegment)) Then
            If (.Key > 0) Then
                .Grouped = True
                Set lyTemp = .Layer
                Call mSelection.Add(graphicObject)
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
    
    nCount = Document.Group.Count

    '
    ' create the feature chains
    '
    Dim GraphicObj() As Esprit.graphicObject
    If mSelection.Count > 0 Then
        GraphicObj = Document.FeatureRecognition.CreateAutoChains(mSelection)
    End If
    

    Call DeleteSmallPointChainFeature(lyTemp.Name, 0.2)
    Document.Refresh
    
    '
    ' loop through all of the newly created features and reverse the closed ones
    '
    Dim Fc As Esprit.FeatureChain
    For i = LBound(GraphicObj) To UBound(GraphicObj)
        If GraphicObj(i).GraphicObjectType = espFeatureChain Then
            Set Fc = GraphicObj(i)
            'If FC.IsClosed Then FC.Reverse
            'Fc.Reverse
            If minValue(Fc.Extremity(espExtremityStart).x, Fc.Extremity(espExtremityEnd).x) = Fc.Extremity(espExtremityEnd).x Then Fc.Reverse
        End If
    Next
    
    Fc.Grouped = True
    Document.Refresh
    
    
    'Count FC
    Dim nCnt As Integer
    nCnt = 0
    For Each fcCnt In Document.FeatureChains
        If (fcCnt.Layer.Name = lyTemp.Name) Then
            nCnt = nCnt + 1
        End If
    Next
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Keep multi generated FCs and exit Function
    If nCnt > 1 Then
        Step2_6 = -991
        Exit Function
    End If
    
'    If (UBound(GraphicObj) >= 1) Then
'        'Call MsgBox("More than 2 Chain features are made. Please check it.", vbCritical, "Error in Front Truning.")
'        Fc.Grouped = True
'        Document.Refresh
'        Step2_6 = -991
'        Exit Function
'    End If
    
    Dim FileName As String
    FileName = "2_FRONT_TURNING.prc"

'    Dim feature As Esprit.FeatureChain
'    Set feature = Document.FeatureChains.Item(1)
    Dim M_TechnologyUtility As EspritTechnology.TechnologyUtility
    Set M_TechnologyUtility = Document.TechnologyUtility
    Dim tech() As EspritTechnology.Technology
    tech = M_TechnologyUtility.OpenProcess(EspritUserFolder & FileName)
    
    
    Dim Op As Esprit.Operation
    
    Set Op = Document.Operations.Add(tech(0), Fc)
    Op.Name = "2-1. FRONT TURNING"
'    Call SetCustomLong(Op.CustomProperties, "SortOrder", 21)
'    Call DeleteDummyOperation
    
    'Very Important - destroy all of the Objects
    'Set M_Segment = Nothing
    'Set doc = Nothing
    'Set App = Nothing
    
    Step2_6 = 1 'Success Code
End Function


Private Function CopyPlane(Optional SourcePlane As Esprit.Plane, Optional DestinationPlane As Esprit.Plane) As Esprit.Plane

    If SourcePlane Is Nothing Then

        Set SourcePlane = Document.ActivePlane

    End If

    If DestinationPlane Is Nothing Then

        Set DestinationPlane = GetPlane("Copy of " & SourcePlane.Name)

    End If

    With DestinationPlane

        .IsView = SourcePlane.IsView

        .IsWork = SourcePlane.IsWork

        .x = SourcePlane.x

        .y = SourcePlane.y

        .Z = SourcePlane.Z

        .Ux = SourcePlane.Ux

        .Uy = SourcePlane.Uy

        .Uz = SourcePlane.Uz

        .Vx = SourcePlane.Vx

        .Vy = SourcePlane.Vy

        .Vz = SourcePlane.Vz

        .Wx = SourcePlane.Wx

        .Wy = SourcePlane.Wy

        .Wz = SourcePlane.Wz

    End With

    Set CopyPlane = DestinationPlane

End Function

 

Private Sub FindCrossSectionAreas()

    '

    ' prompt the user to pick a solid

    '

    Dim sl As Esprit.Solid

    On Error Resume Next

    Set sl = Document.GetAnyElement("Select Reference Solid", espSolidModel)

    If sl Is Nothing Then Exit Sub ' in case user presses escape

    On Error GoTo 0 ' to disable further error trapping

    '

    ' prompt the user for the top, bottom, and increment, with appropriate defaults

    '

    Dim ZTop As Double, ZBottom As Double, ZIncrement As Double

    If Document.SystemUnit = espInch Then

        ZTop = 1

        Application.Configuration.ConfigurationFeatureRecognition.Tolerance = 0.001

    Else

        ZTop = 25

        Application.Configuration.ConfigurationFeatureRecognition.Tolerance = 0.025

    End If

    ZTop = Val(InputBox("Input Start Section Top Z Value", "Top Z", ZTop))

    Do

        ZBottom = Val(InputBox("Input End Section Bottom Z Value", "Bottom Z", 0))

    Loop Until (ZBottom <> ZTop)

    ZIncrement = (ZTop - ZBottom) / 10

    Do

        ZIncrement = Val(InputBox("Input Increment Z Value", "Increment Z", ZIncrement))

    Loop Until (ZIncrement <> 0)

    ZIncrement = Abs(ZIncrement) * Sgn(ZTop - ZBottom) ' make sure sign is correct

    '

    ' make a plane parallel to current active plane

    '

    Dim PL(0 To 1) As Esprit.Plane

    Set PL(0) = Document.ActivePlane

    Set PL(1) = CopyPlane(PL(0), GetPlane("Parallel to " & PL(0).Name))

    PL(1).Activate

    Call PL(1).Parallel(0, 0, ZTop)

    '

    ' add the solid to the selection set

    '

    Dim SS As Esprit.SelectionSet

    Set SS = GetSelectionSet("Temp")

    SS.RemoveAll

    Call SS.Add(sl)

    '

    ' prepare to display the results

    '

    Application.OutputWindow.Clear

    Application.OutputWindow.Visible = True

    '

    ' loop through all of the Z levels

    '

    Dim Z As Double, DepthLevel As Long

    Z = ZTop

    DepthLevel = 1

    Do Until ((Z < ZBottom) And (ZBottom <= ZTop)) Or ((Z > ZBottom) And (ZBottom >= ZTop))

        '

        ' create the cross sections

        '

        Dim Sections() As Esprit.graphicObject

        Sections = Document.FeatureRecognition.CreatePartProfileCrossSection(SS, PL(1), espFeatureChains, False)

        '

        ' loop through the resulting graphic objects

        '

        Dim i As Long, SectionNumber As Long

        SectionNumber = 1

        For i = LBound(Sections) To UBound(Sections)

            If Sections(i).GraphicObjectType = espFeatureChain Then

                Dim Fc As Esprit.FeatureChain

                Set Fc = Sections(i)

                '

                ' display the results

                '

                With Application.OutputWindow

                    Call .Text("The Area of Section " & SectionNumber & " at Depth " & DepthLevel & " (Z=" & Round(Z, 5) & ") is " & Round(Fc.Area, 5) & vbCrLf)

                End With

                SectionNumber = SectionNumber + 1

            End If

        Next

        '

        ' update for next section

        '

        Z = Z - ZIncrement

        DepthLevel = DepthLevel + 1

        Call PL(1).Parallel(0, 0, (-1) * ZIncrement)

    Loop

    '

    ' return to previous plane

    '

    PL(0).Activate

End Sub



Private Function GetPlane(PlaneName As String) As Esprit.Plane

    On Error Resume Next

    Set GetPlane = Document.Planes.Item(PlaneName)

    On Error GoTo 0

    If GetPlane Is Nothing Then

        Set GetPlane = Document.Planes.Add(PlaneName)

    End If

End Function

Private Function SetPlaneVectors(PlaneObj As Esprit.Plane, x As Double, y As Double, Z As Double, Ux As Double, Uy As Double, Uz As Double, Vx As Double, Vy As Double, Vz As Double, Wx As Double, Wy As Double, Wz As Double) As Esprit.Plane

    With PlaneObj

        .x = x

        .y = y

        .Z = Z

        .Ux = Ux

        .Uy = Uy

        .Uz = Uz

        .Vx = Vx

        .Vy = Vy

        .Vz = Vz

        .Wx = Wx

        .Wy = Wy

        .Wz = Wz

    End With

    Set SetPlaneVectors = PlaneObj

End Function

 



Private Function GetSelectionSet(SSName As String) As Esprit.SelectionSet

    On Error Resume Next

    Set GetSelectionSet = Document.SelectionSets.Item(SSName)

    On Error GoTo 0

    If GetSelectionSet Is Nothing Then

        Set GetSelectionSet = Document.SelectionSets.Add(SSName)

    End If

End Function

Private Sub ScanOperations()

    ' for any operations which are not in the library yet

    ' the following line of code is necessary to skip past them

    On Error Resume Next

    Call Application.OutputWindow.Clear

    Application.OutputWindow.Visible = True

    Application.OutputWindow.Clear

    Dim Op As Esprit.Operation

    ' must use generic Technology object from EspritTechnology library to find TypeName

    Dim OpTech As EspritTechnology.Technology

    For i = 1 To Document.Operations.Count

        Set Op = Document.Operations.Item(i)

        Set OpTech = Op.Technology

        ' Strip the first four characters (Tech) off of the technology typename for display

        Call Application.OutputWindow.Text("OP" & Op.Key & " is a " & Op.Name & vbCrLf)
        Call Application.OutputWindow.Text("OP" & Op.Key & " is a " & Op.TypeName & vbCrLf)
        
        Call Application.OutputWindow.Text("OP" & Op.Key & " is a " & OpTech.TechnologyType & vbCrLf)
        Call Application.OutputWindow.Text("OP" & Op.Key & " is a " & OpTech.Name & vbCrLf)

        Call Application.OutputWindow.Text("OP" & Op.Key & " is on Layer " & Op.Layer.Name & vbCrLf)

        If Document.SystemUnit = espMetric Then

            Call Application.OutputWindow.Text("OP" & Op.Key & " Length of Feed is " & Op.LengthOfFeed & " MM" & vbCrLf)

            Call Application.OutputWindow.Text("OP" & Op.Key & " Length of Rapid is " & Op.LengthOfRapid & " MM" & vbCrLf)

        Else

            Call Application.OutputWindow.Text("OP" & Op.Key & " Length of Feed is " & Op.LengthOfFeed & " IN" & vbCrLf)

            Call Application.OutputWindow.Text("OP" & Op.Key & " Length of Rapid is " & Op.LengthOfRapid & " IN" & vbCrLf)

        End If

        ' This outputs a blank line between operations

        Call Application.OutputWindow.Text(vbCrLf)

    Next

End Sub


'Sub Step3()
'    Call generateSolidmilTurn("0DEG", "ROUGH ENDMILL R6.0", "1")
'    Call generateSolidmilTurn("120DEG", "ROUGH ENDMILL R6.0", "2")
'    Call generateSolidmilTurn("240DEG", "ROUGH ENDMILL R6.0", "3")
'    Call DeleteDummyOperation
'
'End Sub
Function generateSolidmilTurn(strWorkPlaneName As String, strBaseLayerName As String, strOperationOrder As String, Optional ByVal bTolerance As Double = 0.1) As Integer
'3> 솔리드 밀턴 Tool Path 생성
'1)  Layer를 2. Rough Endmill R6.0 으로 선택
'2)  0DEG 선택(기본)
'3)  STL 객체의 Part Profile 로 외곽선을 생성 (공차 0.1, 너무 작으면 자동 연결시 분할됨)
'4)  CUT-OFF layer상의 좌측 기준선과 STL 객체의 접점을 생성
'접점상, 접점하 2개가 나올 것
'5)  상측작업
'접점 기준으로 좌측의 모든 요소의 Y최고점을 확인
'X축 기준으로 정렬했을 때
'접점상X 값 - 1 보다 뒤에 있는 최고점(Y) > 접점상Y + 1
'Yes)
'접점상 길이를 따라 2 위치의 진입경계점(상)을 구하고
'해당 점에서 90도로 길이 6 의 선분 생성
'    No)
'    진입경계점상 우측: 진입경계점상Y ~ 0 Y 사이 trim
'진입경계점상 좌측: Part Profile 로 생성한 외곽선 내측 trim

    generateSolidmilTurn = 0

    'Tolerance set to 0.1 (default 공차)
    Dim bOriTolerance As Double
    bOriTolerance = Application.Configuration.ConfigurationFeatureRecognition.Tolerance
    If Document.SystemUnit = espInch Then
        Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bTolerance * 3.9
    Else
        Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bTolerance
    End If

'1)  Layer를 2. Rough Endmill R6.0 으로 선택
    Dim lyOri As Esprit.Layer
    Dim strOriginalLayer As String
    strOriginalLayer = strBaseLayerName
    
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

'3)  STL 객체의 Part Profile 로 외곽선을 생성 (공차 0.1, 너무 작으면 자동 연결시 분할됨)

    Dim nYTop As Double
    Dim nYBottom As Double
    
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item(strWorkPlane)
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add(strWorkPlane)
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
    
    If GetPartProfile(lyTemp, strWorkPlane, "STL", bTolerance) = -1 Then
        Call MsgBox("Cannot find any STL model in STL Layer. Please load an STL file and try it again.")
        Exit Function
    End If
    
    For Each goRef In Esprit.Document.GraphicsCollection
        With goRef
        If (.Layer.Name = lyTemp.Name And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType = espArc Or .GraphicObjectType = espPoint Or .GraphicObjectType = espSegment)) Then
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
    'Document.Refresh
    
    Call mSelection.ChangeLayer(Document.ActiveLayer, 0)
    

'4) 상측작업
    'CUT-OFF layer상의 좌측 기준선(원본),(확장)
    Dim sgLeftSegmentCutOffOriginal As Esprit.Segment
    Dim sgLeftSegmentCutOffExtended As Esprit.Segment
    
    'CUT-OFF layer상의 좌측 기준선에서 X-1 선분 (기준선[X-1])
    Dim sgLeftSegmentCutOffLeft1mm As Esprit.Segment
    
    'CheckSegment(X-1,Y+1) / 접점에서 1차 확인선
    Dim sgLeftSegmentCutOffYExt1 As Esprit.Segment
    
    '접점상,하
'    Dim ptTop As Esprit.Point
'    Dim ptBottom As Esprit.Point
    
    '진입점상, 하
    Dim ptTopL1 As Esprit.Point
    Dim ptBottomL1 As Esprit.Point
    
    
    Dim sgSegment As Esprit.Segment
    Dim sgArc As Esprit.Arc
    
    Dim nSegKey As Integer
    Dim nArcKey As Integer
    Dim nptTopL1Key As Integer
    Dim nptBottomL1Key As Integer
    
    Dim ptArc(3) As Point
    
    Dim intersections() As EspritGeometryBase.IComPoint 'EspritGeometry.ComPoint
    
   ' Set ptTop = Document.Points(16) '("P16")
    'Set ptBottom = Document.Points(14) '("P17")
    'Set sgSegment = Document.Segments.Item("122")
   ' sgArc.XStart
    
    'CUT-OFF layer상의 좌측 기준선(원본)
    Set sgLeftSegmentCutOffOriginal = GetLeftSegmentFromConnectionCuffOff()
    
    'CUT-OFF layer상의 좌측 기준선(확장)
    Set sgLeftSegmentCutOffExtended = Document.GetSegment(Document.GetPoint(sgLeftSegmentCutOffOriginal.XStart, maxValue(sgLeftSegmentCutOffOriginal.YStart, sgLeftSegmentCutOffOriginal.YEnd), sgLeftSegmentCutOffOriginal.ZStart), _
                                                            Document.GetPoint(sgLeftSegmentCutOffOriginal.XEnd, maxValue(sgLeftSegmentCutOffOriginal.YStart, sgLeftSegmentCutOffOriginal.YEnd) * (-1), sgLeftSegmentCutOffOriginal.ZStart))
    'Call Document.Segments.Add(Document.GetPoint(sgLeftSegmentCutOffOriginal.XStart, maxValue(sgLeftSegmentCutOffOriginal.YStart, sgLeftSegmentCutOffOriginal.YEnd), sgLeftSegmentCutOffOriginal.ZStart), _
    '                                                        Document.GetPoint(sgLeftSegmentCutOffOriginal.XEnd, maxValue(sgLeftSegmentCutOffOriginal.YStart, sgLeftSegmentCutOffOriginal.YEnd) * (-1), sgLeftSegmentCutOffOriginal.ZStart))
    
    '기준선[X-1]
    Set sgLeftSegmentCutOffLeft1mm = Document.GetSegment(Document.GetPoint(sgLeftSegmentCutOffOriginal.XStart - 1, maxValue(sgLeftSegmentCutOffOriginal.YStart, sgLeftSegmentCutOffOriginal.YEnd), sgLeftSegmentCutOffOriginal.ZStart), _
                                                            Document.GetPoint(sgLeftSegmentCutOffOriginal.XEnd - 1, maxValue(sgLeftSegmentCutOffOriginal.YStart, sgLeftSegmentCutOffOriginal.YEnd) * (-1), sgLeftSegmentCutOffOriginal.ZStart))
    'Call Document.Segments.Add(Document.GetPoint(sgLeftSegmentCutOffOriginal.XStart - 1, maxValue(sgLeftSegmentCutOffOriginal.YStart, sgLeftSegmentCutOffOriginal.YEnd), sgLeftSegmentCutOffOriginal.ZStart), _
    '                                                        Document.GetPoint(sgLeftSegmentCutOffOriginal.XEnd - 1, maxValue(sgLeftSegmentCutOffOriginal.YStart, sgLeftSegmentCutOffOriginal.YEnd) * (-1), sgLeftSegmentCutOffOriginal.ZStart))
    
    
    'sgLeftSegmentCutOff.Grouped = True
    For Each sgSegmentOn In Esprit.Document.Segments
            If (sgSegmentOn.Layer.Number = Document.ActiveLayer.Number) Then
                If (maxValue(sgSegmentOn.YStart, sgSegmentOn.YEnd) > 0 _
                And minValue(sgSegmentOn.XStart, sgSegmentOn.XEnd) <= sgLeftSegmentCutOffExtended.XStart _
                And maxValue(sgSegmentOn.XStart, sgSegmentOn.XEnd) > sgLeftSegmentCutOffExtended.XStart) Then
                    Set sgSegment = sgSegmentOn
                    nSegKey = sgSegment.Key
                    
                    Exit For
                End If
            End If
    Next
    
    If sgSegment Is Nothing Then
        For Each sgArcOn In Esprit.Document.Arcs
            If (sgArcOn.Layer.Number = Document.ActiveLayer.Number) Then
                If (maxValue(sgArcOn.Extremity(espExtremityStart).y, sgArcOn.Extremity(espExtremityEnd).y) > 0 _
                And minValue(sgArcOn.Extremity(espExtremityStart).x, sgArcOn.Extremity(espExtremityEnd).x) <= sgLeftSegmentCutOffExtended.XStart _
                And maxValue(sgArcOn.Extremity(espExtremityStart).x, sgArcOn.Extremity(espExtremityEnd).x) > sgLeftSegmentCutOffExtended.XStart) Then
                    Set sgArc = sgArcOn
                    nArcKey = sgArc.Key
                    Exit For
                End If
            End If
        Next
    Else
        sgSegment.Grouped = True
'        sgSegment.Grouped = False
    End If
    
    If sgArc Is Nothing Then
    'It must have intersection point (sothing wrong case)
    Else
        sgArc.Grouped = True
'        sgArc.Grouped = False
    End If
    
    'sgLeftSegmentCutOffExtended.Grouped = True
    
    '접점상 ptTop
    If Not (sgSegment Is Nothing) Then
        Set ptTop = IntersectArcsSegments(sgLeftSegmentCutOffExtended, sgSegment)
    ElseIf Not (sgArc Is Nothing) Then
        Set ptTop = IntersectArcsSegments(sgLeftSegmentCutOffExtended, sgArc)
    Else
        Call MsgBox("Cannot find any connection point. Please check it, and try it again.")
        Exit Function
    End If
    'ptTop.Grouped = True
    
    Set sgLeftSegmentCutOffYExt1 = Document.GetSegment(Document.GetPoint(sgLeftSegmentCutOffOriginal.XStart - 1, ptTop.y + 1, sgLeftSegmentCutOffOriginal.ZStart), _
                                                            Document.GetPoint(sgLeftSegmentCutOffOriginal.XEnd, ptTop.y + 1, sgLeftSegmentCutOffOriginal.ZStart))
'    Set sgLeftSegmentCutOffYExt1 = Document.Segments.Add(Document.GetPoint(sgLeftSegmentCutOffOriginal.XStart - 1, ptTop.Y + 1, sgLeftSegmentCutOffOriginal.ZStart), _
'                                                            Document.GetPoint(sgLeftSegmentCutOffOriginal.XEnd, ptTop.Y + 1, sgLeftSegmentCutOffOriginal.ZStart))

    
    'Document.Refresh
    
'    Set sgSegment = Nothing
    If Not (sgSegment Is Nothing) Then
        Set sgSegment = Nothing
    End If
    If Not (sgArc Is Nothing) Then
        Set sgArc = Nothing
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' 1) Is get an intersect point with Y+1?
    'sgLeftSegmentCutOff.Grouped = True
    For Each sgSegmentOn In Esprit.Document.Segments
            If (sgSegmentOn.Layer.Number = Document.ActiveLayer.Number) Then
                If (maxValue(sgSegmentOn.YStart, sgSegmentOn.YEnd) > 0 _
                And minValue(sgSegmentOn.XStart, sgSegmentOn.XEnd) < maxValue(sgLeftSegmentCutOffYExt1.XStart, sgLeftSegmentCutOffYExt1.XEnd) _
                And maxValue(sgSegmentOn.XStart, sgSegmentOn.XEnd) > minValue(sgLeftSegmentCutOffYExt1.XStart, sgLeftSegmentCutOffYExt1.XEnd) _
                And minValue(sgSegmentOn.YStart, sgSegmentOn.YEnd) <= sgLeftSegmentCutOffYExt1.YStart _
                And maxValue(sgSegmentOn.YStart, sgSegmentOn.YEnd) > sgLeftSegmentCutOffYExt1.YStart) Then
                    Set sgSegment = sgSegmentOn
                    sgSegment.Grouped = True
                    Exit For
                End If
            End If
    Next
    
    If sgSegment Is Nothing Then
        For Each sgArcOn In Esprit.Document.Arcs
            If (sgArcOn.Layer.Number = Document.ActiveLayer.Number) Then
                If (maxValue(sgArcOn.Extremity(espExtremityStart).y, sgArcOn.Extremity(espExtremityEnd).y) > 0 _
                And minValue(sgArcOn.Extremity(espExtremityStart).x, sgArcOn.Extremity(espExtremityEnd).x) < maxValue(sgLeftSegmentCutOffYExt1.XStart, sgLeftSegmentCutOffYExt1.XEnd) _
                And maxValue(sgArcOn.Extremity(espExtremityStart).x, sgArcOn.Extremity(espExtremityEnd).x) > minValue(sgLeftSegmentCutOffYExt1.XStart, sgLeftSegmentCutOffYExt1.XEnd) _
                And minValue(sgArcOn.Extremity(espExtremityStart).y, sgArcOn.Extremity(espExtremityEnd).y) <= sgLeftSegmentCutOffYExt1.YStart _
                And maxValue(sgArcOn.Extremity(espExtremityStart).y, sgArcOn.Extremity(espExtremityEnd).y) > sgLeftSegmentCutOffYExt1.YStart) Then
                    Set sgArc = sgArcOn
                    sgArc.Grouped = True
                    Exit For
                End If
            End If
        Next
    Else
        sgSegment.Grouped = True
'        sgSegment.Grouped = False
    End If
    
    Dim sgArcCom As ComArc
    
    
    '교차선/호 trim
    If Not (sgSegment Is Nothing) Then
            Set ptTopL1 = IntersectArcsSegments(sgSegment, sgLeftSegmentCutOffYExt1)
            If Not (ptTopL1 Is Nothing) Then
                If (sgSegment.XStart > sgSegment.XEnd) Then
                    sgSegment.Reverse
                End If
                sgSegment.Grouped = True
                sgSegment.XEnd = ptTopL1.x
                sgSegment.YEnd = ptTopL1.y
                sgSegment.ZEnd = ptTopL1.Z
            End If
    ElseIf Not (sgArc Is Nothing) Then
            Set ptTopL1 = IntersectArcsSegments(sgArc, sgLeftSegmentCutOffYExt1)
            If Not (ptTopL1 Is Nothing) Then
                If (sgArc.Extremity(espExtremityStart).x > sgArc.Extremity(espExtremityEnd).x) Then
                    sgArc.Reverse
                End If
                
                Set ptArc(0) = sgArc.Extremity(espExtremityStart)
                Set ptArc(1) = sgArc.PointAlong(sgArc.LengthAlong(ptTopL1) / 2)
                Set ptArc(2) = ptTopL1
                
                Set sgArcCom = GetGeoUtility().Arc3(PointToComPoint(ptArc(0)), PointToComPoint(ptArc(0)), _
                                                    PointToComPoint(ptArc(1)), PointToComPoint(ptArc(1)), _
                                                    PointToComPoint(ptArc(2)), PointToComPoint(ptArc(2)))
                
                Call Document.GraphicsCollection.Remove(sgArc.GraphicsCollectionIndex)
                Set sgArc = ComArcToArc(sgArcCom)
            End If
    End If
    
    If Not (ptTopL1 Is Nothing) Then
        nptTopL1Key = ptTopL1.Key
        'ptTopL1.Grouped = True
    Else
        Set sgSegment = Nothing
        Set sgArc = Nothing
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' 2) Is nothing get an intersect point with Y+1?
    'sgLeftSegmentCutOff.Grouped = True
        For Each sgSegmentOn In Esprit.Document.Segments
                If (sgSegmentOn.Layer.Number = Document.ActiveLayer.Number And sgSegmentOn.Key <> sgLeftSegmentCutOffYExt1.Key) Then
                    If (maxValue(sgSegmentOn.YStart, sgSegmentOn.YEnd) > 0 _
                    And minValue(sgSegmentOn.XStart, sgSegmentOn.XEnd) <= sgLeftSegmentCutOffLeft1mm.XStart _
                    And maxValue(sgSegmentOn.XStart, sgSegmentOn.XEnd) > sgLeftSegmentCutOffLeft1mm.XStart) Then
                        Set sgSegment = sgSegmentOn
                        sgSegment.Grouped = True
                        Exit For
                    End If
                End If
        Next
        
        For Each sgArcOn In Esprit.Document.Arcs
            If (sgArcOn.Layer.Number = Document.ActiveLayer.Number) Then
                If (maxValue(sgArcOn.Extremity(espExtremityStart).y, sgArcOn.Extremity(espExtremityEnd).y) > 0 _
                And minValue(sgArcOn.Extremity(espExtremityStart).x, sgArcOn.Extremity(espExtremityEnd).x) <= sgLeftSegmentCutOffLeft1mm.XStart _
                And maxValue(sgArcOn.Extremity(espExtremityStart).x, sgArcOn.Extremity(espExtremityEnd).x) > sgLeftSegmentCutOffLeft1mm.XStart) Then
                    Set sgArc = sgArcOn
                    'For Debug
                    'sgArc.Grouped = True
                    Exit For
                End If
            End If
        Next
        '교차선/호 trim
        If Not (sgSegment Is Nothing) Then
                Set ptTopL1 = IntersectArcsSegments(sgSegment, sgLeftSegmentCutOffLeft1mm)
            If Not (ptTopL1 Is Nothing) Then
                If (sgSegment.XStart > sgSegment.XEnd) Then
                    sgSegment.Reverse
                End If
                sgSegment.Grouped = True
                
                sgSegment.XEnd = ptTopL1.x
                sgSegment.YEnd = ptTopL1.y
                sgSegment.ZEnd = ptTopL1.Z
            End If
        ElseIf Not (sgArc Is Nothing) Then
                Set ptTopL1 = IntersectArcsSegments(sgArc, sgLeftSegmentCutOffLeft1mm)
            If Not (ptTopL1 Is Nothing) Then
                If (sgArc.Extremity(espExtremityStart).x > sgArc.Extremity(espExtremityEnd).x) Then
                    sgArc.Reverse
                End If
                
                Set ptArc(0) = sgArc.Extremity(espExtremityStart)
                Set ptArc(1) = sgArc.PointAlong(sgArc.LengthAlong(ptTopL1) / 2)
                Set ptArc(2) = ptTopL1
                
                Set sgArcCom = GetGeoUtility().Arc3(PointToComPoint(ptArc(0)), PointToComPoint(ptArc(0)), _
                                                    PointToComPoint(ptArc(1)), PointToComPoint(ptArc(1)), _
                                                    PointToComPoint(ptArc(2)), PointToComPoint(ptArc(2)))
                
                Call Document.GraphicsCollection.Remove(sgArc.GraphicsCollectionIndex)
                Set sgArc = ComArcToArc(sgArcCom)
                sgArc.Grouped = True
                'Call sgArc.Extremity(espExtremityEnd).SetXyz(ptTopL1.X, ptTopL1.Y, ptTopL1.Z)
            End If
        End If
    End If
    
'진입경계점 (상)
    nptTopL1Key = ptTopL1.Key
    'ptTopL1.Grouped = True
    
'진입유도선(상)
    Dim segTop1 As Esprit.Segment
    '해당 점에서 45도로 길이 6 의 선분 생성
    'Set segTop1 = Document.Segments.Add(Document.GetPoint(ptTopL1.X + Sqr(6), ptTopL1.Y + Sqr(6), ptTopL1.Z), ptTopL1)
    '해당 점에서 90도로 길이 6 의 선분 생성 requested by Andrew at 03.29.2018
    Set segTop1 = Document.Segments.Add(Document.GetPoint(ptTopL1.x, ptTopL1.y + 6, ptTopL1.Z), ptTopL1)
    
    'Debug
    'Document.Refresh
    
    If (segTop1.XStart > segTop1.XEnd) Then
        segTop1.Reverse
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''
    'Trim below the extended segment
    For Each goObject In Esprit.Document.GraphicsCollection
    With goObject
    If (.Layer.Name = Document.ActiveLayer.Name) Then
        If (.GraphicObjectType = espSegment) Then
            Set tmpSegment = goObject
            tmpSegment.Grouped = True
            With tmpSegment
'                If (.Key <> segTop1.Key And (minValue(.XStart, .XEnd) >= segTop1.XStart) _
'                    And (minValue(.YStart, .YEnd) >= 0) And (minValue(.YStart, .YEnd) < segTop1.YEnd)) Then
                If (.Key <> segTop1.Key And (minValue(.XStart, .XEnd) >= segTop1.XStart) _
                    And (minValue(.YStart, .YEnd) >= 0)) Then
                    
                    Call Document.GraphicsCollection.Remove(.GraphicsCollectionIndex)
                    'Document.Refresh
                End If
'                If (.Key <> segTop1.Key And (minValue(.YStart, .YEnd) >= 0) And (maxValue(.XStart, .XEnd) >= segTop1.XEnd)) Then
'                    Call Document.GraphicsCollection.Remove(.GraphicsCollectionIndex)
'                    Document.Refresh
'                End If
            
            End With
            tmpSegment.Grouped = False
        ElseIf (.GraphicObjectType = espArc) Then
            Set tmpArc = goObject
            tmpArc.Grouped = True
            
            With tmpArc
'                If (.Key <> segTop1.Key And (minValue(.Extremity(espExtremityStart).X, .Extremity(espExtremityEnd).X) >= segTop1.XStart) _
'                    And (maxValue(.Extremity(espExtremityStart).Y, .Extremity(espExtremityEnd).Y) >= 0) _
'                    And (minValue(.Extremity(espExtremityStart).Y, .Extremity(espExtremityEnd).Y) < segTop1.YEnd)) Then
                If (.Key <> segTop1.Key And (minValue(.Extremity(espExtremityStart).x, .Extremity(espExtremityEnd).x) >= segTop1.XStart) _
                    And (maxValue(.Extremity(espExtremityStart).y, .Extremity(espExtremityEnd).y) >= 0)) Then
                    Call Document.GraphicsCollection.Remove(.GraphicsCollectionIndex)
                    'Document.Refresh
                End If
'                If (.Key <> segTop1.Key And (maxValue(.Extremity(espExtremityStart).Y, .Extremity(espExtremityEnd).Y) >= 0) _
'                    And (maxValue(.Extremity(espExtremityStart).X, .Extremity(espExtremityEnd).X) >= segTop1.XEnd)) Then
'                    Call Document.GraphicsCollection.Remove(.GraphicsCollectionIndex)
'                    Document.Refresh
'                End If
            
            End With
            tmpArc.Grouped = False
        
        Else
'                .Grouped = False
        End If
    ElseIf (.GraphicObjectType <> espUnknown) Then
        .Grouped = False
    End If
    End With
    Next
    
    'Debug
    'Document.Refresh
'''''''''''''''''''''''''''''''''''''''''''
'5) 하측작업
    'CUT-OFF layer상의 좌측 기준선(원본),(확장)
    'Dim sgLeftSegmentCutOffOriginal As Esprit.Segment
    'Dim sgLeftSegmentCutOffExtended As Esprit.Segment
    
    If sgLeftSegmentCutOffExtended.YStart > 0 Then
        sgLeftSegmentCutOffExtended.YStart = 0
    Else
        sgLeftSegmentCutOffExtended.YEnd = 0
    End If
    'CUT-OFF layer상의 좌측 기준선에서 X-1 선분 (기준선[X-1])
    'Dim sgLeftSegmentCutOffLeft1mm As Esprit.Segment
    
    'CheckSegment(X-1,Y-1) for 접점(하) / 접점에서 1차 확인선
    'Dim sgLeftSegmentCutOffYExt1 As Esprit.Segment
    
    '접점상,하
    'Dim poIntersect  As Esprit.Point
    'Dim ptBottom As Esprit.Point
    Dim ptIntersect As Esprit.Point
    
    
    '진입점상, 하
    'Dim poIntersect L1 As Esprit.Point
    'Dim ptBottomL1 As Esprit.Point
    
    
    Set sgSegment = Nothing
    Set sgArc = Nothing
    
    'Dim nSegKey As Integer
    'Dim nArcKey As Integer
    'Dim npoIntersect L1Key As Integer
    'Dim nptBottomL1Key As Integer
    
    'Dim intersections() As EspritGeometryBase.IComPoint 'EspritGeometry.ComPoint
    
    '기준선[X-1]
    'sgLeftSegmentCutOffLeft1mm
    
    'sgLeftSegmentCutOff.Grouped = True
    sgLeftSegmentCutOffExtended.Grouped = True
    
    For Each sgSegmentOn In Esprit.Document.Segments
            If (sgSegmentOn.Layer.Number = Document.ActiveLayer.Number) Then
                If (minValue(sgSegmentOn.YStart, sgSegmentOn.YEnd) < 0 _
                And minValue(sgSegmentOn.XStart, sgSegmentOn.XEnd) < sgLeftSegmentCutOffExtended.XStart _
                And maxValue(sgSegmentOn.XStart, sgSegmentOn.XEnd) > sgLeftSegmentCutOffExtended.XStart) Then
                    Set sgSegment = sgSegmentOn
                    nSegKey = sgSegment.Key
                    
                    Exit For
                End If
            End If
    Next
    
    If sgSegment Is Nothing Then
        For Each sgArcOn In Esprit.Document.Arcs
            If (sgArcOn.Layer.Number = Document.ActiveLayer.Number) Then
                If (minValue(sgArcOn.Extremity(espExtremityStart).y, sgArcOn.Extremity(espExtremityEnd).y) < 0 _
                And minValue(sgArcOn.Extremity(espExtremityStart).x, sgArcOn.Extremity(espExtremityEnd).x) < sgLeftSegmentCutOffExtended.XStart _
                And maxValue(sgArcOn.Extremity(espExtremityStart).x, sgArcOn.Extremity(espExtremityEnd).x) > sgLeftSegmentCutOffExtended.XStart) Then
                    Set sgArc = sgArcOn
                    nArcKey = sgArc.Key
                    Exit For
                End If
            End If
        Next
    Else
        sgSegment.Grouped = True
'        sgSegment.Grouped = False
    End If
    
    If sgArc Is Nothing Then
    'It must have intersection point (sothing wrong case)
    Else
        sgArc.Grouped = True
'        sgArc.Grouped = False
    End If
    
    'sgLeftSegmentCutOffExtended.Grouped = True
    '접점하 ptIntersect
    If Not (sgSegment Is Nothing) Then
        Set ptIntersect = IntersectArcsSegments(sgLeftSegmentCutOffExtended, sgSegment)
    ElseIf Not (sgArc Is Nothing) Then
        Set ptIntersect = IntersectArcsSegments(sgLeftSegmentCutOffExtended, sgArc)
    Else
        Call MsgBox("Cannot find any connection point. Please check it, and try it again.")
        Exit Function
    End If
        
    'Debug
    ptIntersect.Grouped = True
    
    Set sgLeftSegmentCutOffYExt1 = Document.GetSegment(Document.GetPoint(sgLeftSegmentCutOffExtended.XStart - 1, ptIntersect.y - 1, sgLeftSegmentCutOffExtended.ZStart), _
                                                            Document.GetPoint(sgLeftSegmentCutOffExtended.XEnd, ptIntersect.y - 1, sgLeftSegmentCutOffExtended.ZStart))

'sgLeftSegmentCutOffYExt1.Grouped
    
    'Document.Refresh
sgLeftSegmentCutOffYExt1.Grouped = True
    
'    Set sgSegment = Nothing
    If Not (sgSegment Is Nothing) Then
        Set sgSegment = Nothing
    End If
    If Not (sgArc Is Nothing) Then
        Set sgArc = Nothing
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' 1) Is get an intersect point with Y-1?
    'sgLeftSegmentCutOff.Grouped = True
    For Each sgSegmentOn In Esprit.Document.Segments
            If (sgSegmentOn.Layer.Number = Document.ActiveLayer.Number) Then
                If (minValue(sgSegmentOn.YStart, sgSegmentOn.YEnd) < 0 _
                And minValue(sgSegmentOn.XStart, sgSegmentOn.XEnd) < maxValue(sgLeftSegmentCutOffYExt1.XStart, sgLeftSegmentCutOffYExt1.XEnd) _
                And maxValue(sgSegmentOn.XStart, sgSegmentOn.XEnd) > minValue(sgLeftSegmentCutOffYExt1.XStart, sgLeftSegmentCutOffYExt1.XEnd) _
                And minValue(sgSegmentOn.YStart, sgSegmentOn.YEnd) <= sgLeftSegmentCutOffYExt1.YStart _
                And maxValue(sgSegmentOn.YStart, sgSegmentOn.YEnd) > sgLeftSegmentCutOffYExt1.YStart) Then
                    Set sgSegment = sgSegmentOn
                    sgSegment.Grouped = True
                    Exit For
                End If
            End If
    Next
    
    If sgSegment Is Nothing Then
        For Each sgArcOn In Esprit.Document.Arcs
            If (sgArcOn.Layer.Number = Document.ActiveLayer.Number) Then
                If (minValue(sgArcOn.Extremity(espExtremityStart).y, sgArcOn.Extremity(espExtremityEnd).y) < 0 _
                And minValue(sgArcOn.Extremity(espExtremityStart).x, sgArcOn.Extremity(espExtremityEnd).x) < maxValue(sgLeftSegmentCutOffYExt1.XStart, sgLeftSegmentCutOffYExt1.XEnd) _
                And maxValue(sgArcOn.Extremity(espExtremityStart).x, sgArcOn.Extremity(espExtremityEnd).x) > minValue(sgLeftSegmentCutOffYExt1.XStart, sgLeftSegmentCutOffYExt1.XEnd) _
                And minValue(sgArcOn.Extremity(espExtremityStart).y, sgArcOn.Extremity(espExtremityEnd).y) <= sgLeftSegmentCutOffYExt1.YStart _
                And maxValue(sgArcOn.Extremity(espExtremityStart).y, sgArcOn.Extremity(espExtremityEnd).y) > sgLeftSegmentCutOffYExt1.YStart) Then
                    Set sgArc = sgArcOn
                    sgArc.Grouped = True
                    Exit For
                End If
            End If
        Next
    Else
        sgSegment.Grouped = True
'        sgSegment.Grouped = False
    End If
    
    '교차선/호 trim
    If Not (sgSegment Is Nothing) Then
        Set ptBottomL1 = IntersectArcsSegments(sgSegment, sgLeftSegmentCutOffYExt1)
        If Not (ptBottomL1 Is Nothing) Then
            If (sgSegment.XStart > sgSegment.XEnd) Then
                sgSegment.Reverse
            End If
            sgSegment.XEnd = ptBottomL1.x
            sgSegment.YEnd = ptBottomL1.y
            sgSegment.ZEnd = ptBottomL1.Z
        End If
    ElseIf Not (sgArc Is Nothing) Then
        Set ptBottomL1 = IntersectArcsSegments(sgArc, sgLeftSegmentCutOffYExt1)
        If Not (ptBottomL1 Is Nothing) Then
            If (sgArc.Extremity(espExtremityStart).x > sgArc.Extremity(espExtremityEnd).x) Then
                sgArc.Reverse
            End If
            
            Set ptArc(0) = sgArc.Extremity(espExtremityStart)
            Set ptArc(1) = sgArc.PointAlong(sgArc.LengthAlong(ptBottomL1) / 2)
            Set ptArc(2) = ptBottomL1
            
            Set sgArcCom = GetGeoUtility().Arc3(PointToComPoint(ptArc(0)), PointToComPoint(ptArc(0)), _
                                                PointToComPoint(ptArc(1)), PointToComPoint(ptArc(1)), _
                                                PointToComPoint(ptArc(2)), PointToComPoint(ptArc(2)))
            
            Call Document.GraphicsCollection.Remove(sgArc.GraphicsCollectionIndex)
            Set sgArc = ComArcToArc(sgArcCom)
        
        End If
    End If
    
    If Not (ptBottomL1 Is Nothing) Then
'    If (ptBottomL1.Y < 0 _
'        And ptBottomL1.X < maxValue(sgLeftSegmentCutOffYExt1.XStart, sgLeftSegmentCutOffYExt1.XEnd) _
'        And ptBottomL1.X > minValue(sgLeftSegmentCutOffYExt1.XStart, sgLeftSegmentCutOffYExt1.XEnd) _
'        And ptBottomL1.Y <= sgLeftSegmentCutOffYExt1.YStart _
'        And ptBottomL1.Y > sgLeftSegmentCutOffYExt1.YStart) Then
        
        nptBottomL1Key = ptBottomL1.Key
        ptBottomL1.Grouped = True
    Else
        'ptBottomL1.Grouped = False
        Set sgSegment = Nothing
        Set sgArc = Nothing
    '   Dim mRefGraphicObject() As Esprit.graphicObject
    '   mRefGraphicObject = .FeatureRecognition.CreatePartProfileShadow(mSelection, Document.Planes.Item("0DEG"), espFeatureChains)
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' 2) Is nothing get an intersect point with Y-1?
    'sgLeftSegmentCutOff.Grouped = True
        For Each sgSegmentOn In Esprit.Document.Segments
                If (sgSegmentOn.Layer.Number = Document.ActiveLayer.Number And sgSegmentOn.Key <> sgLeftSegmentCutOffYExt1.Key) Then
                    If (minValue(sgSegmentOn.YStart, sgSegmentOn.YEnd) < 0 _
                    And minValue(sgSegmentOn.XStart, sgSegmentOn.XEnd) <= sgLeftSegmentCutOffLeft1mm.XStart _
                    And maxValue(sgSegmentOn.XStart, sgSegmentOn.XEnd) > sgLeftSegmentCutOffLeft1mm.XStart) Then
                        Set sgSegment = sgSegmentOn
                        sgSegment.Grouped = True
                        Exit For
                    End If
                End If
        Next
        
        For Each sgArcOn In Esprit.Document.Arcs
            If (sgArcOn.Layer.Number = Document.ActiveLayer.Number) Then
                If (minValue(sgArcOn.Extremity(espExtremityStart).y, sgArcOn.Extremity(espExtremityEnd).y) < 0 _
                And minValue(sgArcOn.Extremity(espExtremityStart).x, sgArcOn.Extremity(espExtremityEnd).x) <= sgLeftSegmentCutOffLeft1mm.XStart _
                And maxValue(sgArcOn.Extremity(espExtremityStart).x, sgArcOn.Extremity(espExtremityEnd).x) > sgLeftSegmentCutOffLeft1mm.XStart) Then
                    Set sgArc = sgArcOn
                    'For Debug
                    'sgArc.Grouped = True
                    Exit For
                End If
            End If
        Next
        '교차선/호 trim
        If Not (sgSegment Is Nothing) Then
            Set ptBottomL1 = IntersectArcsSegments(sgSegment, sgLeftSegmentCutOffLeft1mm)
            If Not (ptBottomL1 Is Nothing) Then
                If (sgSegment.XStart > sgSegment.XEnd) Then
                    sgSegment.Reverse
                End If
                sgSegment.XEnd = ptBottomL1.x
                sgSegment.YEnd = ptBottomL1.y
                sgSegment.ZEnd = ptBottomL1.Z
            End If
        ElseIf Not (sgArc Is Nothing) Then
            Set ptBottomL1 = IntersectArcsSegments(sgArc, sgLeftSegmentCutOffLeft1mm)
            If Not (ptBottomL1 Is Nothing) Then
                If (sgArc.Extremity(espExtremityStart).x > sgArc.Extremity(espExtremityEnd).x) Then
                    sgArc.Reverse
                End If
                Set ptArc(0) = sgArc.Extremity(espExtremityStart)
                Set ptArc(1) = sgArc.PointAlong(sgArc.LengthAlong(ptBottomL1) / 2)
                Set ptArc(2) = ptBottomL1
                
                Set sgArcCom = GetGeoUtility().Arc3(PointToComPoint(ptArc(0)), PointToComPoint(ptArc(0)), _
                                                    PointToComPoint(ptArc(1)), PointToComPoint(ptArc(1)), _
                                                    PointToComPoint(ptArc(2)), PointToComPoint(ptArc(2)))
                
                Call Document.GraphicsCollection.Remove(sgArc.GraphicsCollectionIndex)
                Set sgArc = ComArcToArc(sgArcCom)
            End If
        End If
        
    End If
    
'진입경계점(하)
    'ptBottomL1.Grouped = True

'진입유도선(하)
    Dim segBottom1 As Esprit.Segment
    '해당 점에서 45도로 길이 6 의 선분 생성
    'Set segBottom1 = Document.Segments.Add(Document.GetPoint(ptBottomL1.X + Sqr(6), ptBottomL1.Y - Sqr(6), ptBottomL1.Z), ptBottomL1)
    '해당 점에서 90도로 길이 6 의 선분 생성 requested by Andrew at 03.29.2018
    Set segBottom1 = Document.Segments.Add(Document.GetPoint(ptBottomL1.x, ptBottomL1.y - 6, ptBottomL1.Z), ptBottomL1)
    
    'Debug
    segBottom1.Grouped = True
    
'    Set M1_Segment = doc.Segments("122")
    If (segBottom1.XStart > segBottom1.XEnd) Then
        segBottom1.Reverse
    End If
    
    
    ''''''''''''''''''''''''''''''''''''''''''
    'Trim below the extended segment
    For Each goObject In Esprit.Document.GraphicsCollection
    With goObject
    If (.Layer.Name = Document.ActiveLayer.Name) Then
        If (.GraphicObjectType = espSegment) Then
            Set tmpSegment = goObject
            tmpSegment.Grouped = True
            With tmpSegment
                If (.Key <> segTop1.Key And .Key <> segBottom1.Key _
                    And (minValue(.XStart, .XEnd) >= maxValue(minValue(segTop1.XStart, segTop1.XEnd), minValue(segBottom1.XStart, segBottom1.XEnd)))) Then
                    Call Document.GraphicsCollection.Remove(.GraphicsCollectionIndex)
                    'Document.Refresh
'                ElseIf (.Key <> segTop1.Key And .Key <> segBottom1.Key _
'                    And (minValue(.XStart, .XEnd) >= segBottom1.XStart) _
'                    And (maxValue(.YStart, .YEnd) <= 0) _
'                    And (maxValue(.YStart, .YEnd) > segBottom1.YEnd)) Then
                ElseIf (.Key <> segTop1.Key And .Key <> segBottom1.Key _
                    And (minValue(.XStart, .XEnd) >= segBottom1.XStart) _
                    And (maxValue(.YStart, .YEnd) <= 0)) Then
                    Call Document.GraphicsCollection.Remove(.GraphicsCollectionIndex)
                    'Document.Refresh
                End If
            
            End With
            tmpSegment.Grouped = False
        ElseIf (.GraphicObjectType = espArc) Then
            Set tmpArc = goObject
            tmpArc.Grouped = True
            
            With tmpArc
                If (.Key <> segTop1.Key And .Key <> segBottom1.Key _
                    And (minValue(.Extremity(espExtremityStart).x, .Extremity(espExtremityEnd).x) >= maxValue(minValue(segTop1.XStart, segTop1.XEnd), minValue(segBottom1.XStart, segBottom1.XEnd)))) Then
                    Call Document.GraphicsCollection.Remove(.GraphicsCollectionIndex)
                    'Document.Refresh
                
'                ElseIf (.Key <> segTop1.Key And .Key <> segBottom1.Key _
'                    And (minValue(.Extremity(espExtremityStart).X, .Extremity(espExtremityEnd).X) >= segBottom1.XStart) _
'                    And (maxValue(.Extremity(espExtremityStart).Y, .Extremity(espExtremityEnd).Y) <= 0) _
'                    And (maxValue(.Extremity(espExtremityStart).Y, .Extremity(espExtremityEnd).Y) > segBottom1.YEnd)) Then
                ElseIf (.Key <> segTop1.Key And .Key <> segBottom1.Key _
                    And (minValue(.Extremity(espExtremityStart).x, .Extremity(espExtremityEnd).x) >= segBottom1.XStart) _
                    And (maxValue(.Extremity(espExtremityStart).y, .Extremity(espExtremityEnd).y) <= 0)) Then
                    Call Document.GraphicsCollection.Remove(.GraphicsCollectionIndex)
                    'Document.Refresh
                End If
            End With
            tmpArc.Grouped = False
        
        Else
'                .Grouped = False
        End If
    ElseIf (.GraphicObjectType <> espUnknown) Then
        .Grouped = False
    End If
    End With
    Next
    
    
    If Not (sgSegment Is Nothing) Then sgSegment.Grouped = False
    If Not (sgArc Is Nothing) Then sgArc.Grouped = False
    If Not (ptTopL1 Is Nothing) Then ptTopL1.Grouped = False
    If Not (ptBottomL1 Is Nothing) Then ptBottomL1.Grouped = False
    If Not (ptIntersect Is Nothing) Then ptIntersect.Grouped = False
    
    Document.ActiveLayer.Visible = True
    'Debug
    'Document.Refresh

    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
    End With
    For Each goRef In Esprit.Document.GraphicsCollection
        With goRef
        If (.Layer.Name = lyTemp.Name And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType = espArc Or .GraphicObjectType = espPoint Or .GraphicObjectType = espSegment)) Then
            If (.Key > 0) Then
                .Grouped = True
                Call mSelection.Add(goRef)
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next

    '
    ' create the feature chains
    '
    Dim GraphicObj() As Esprit.graphicObject
    If mSelection.Count > 0 Then
        GraphicObj = Document.FeatureRecognition.CreateAutoChains(mSelection)
    End If

    'DeleteSmallPointChainFeature
    Call DeleteSmallPointChainFeature(lyTemp.Name, 0.2)

    'Count FC
    Dim nCnt As Integer
    nCnt = 0
    For Each fcCnt In Document.FeatureChains
        If (fcCnt.Layer.Name = lyTemp.Name) Then
            nCnt = nCnt + 1
        End If
    Next
    
    generateSolidmilTurn = nCnt
    Dim strNewRoughLayer As String
    Dim lyNew As Esprit.Layer
    Dim Fc As Esprit.FeatureChain
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Keep multi generated FCs and exit Function
    If nCnt > 1 Then
        Call MsgBox("STL Profile makes Chain Feature more than 1. Please check it.")
        
        strNewRoughLayer = "[" + strWorkPlane + "]" + strBaseLayerName
        For Each ly In Document.Layers
            If (ly.Name = strNewRoughLayer) Then
                Call Document.Layers.Remove(strNewRoughLayer)
            End If
        Next
        Set lyNew = Document.Layers.Add(strNewRoughLayer)
        
        
        With mSelection
            .RemoveAll
            .AddCopiesToSelectionSet = False
        End With
        For Each goRef In Esprit.Document.GraphicsCollection
            With goRef
            If (.Layer.Name = lyTemp.Name And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType = espArc Or .GraphicObjectType = espPoint Or .GraphicObjectType = espSegment)) Then
                If (.Key > 0) Then
                    .Grouped = True
                    .Layer = lyNew
                    Call mSelection.Add(goRef)
                End If
            ElseIf (.GraphicObjectType <> espUnknown) Then
                .Grouped = False
            End If
            End With
        Next
        For i = LBound(GraphicObj) To UBound(GraphicObj)
            If GraphicObj(i).GraphicObjectType = espFeatureChain Then
                Set Fc = GraphicObj(i)
                'If FC.IsClosed Then FC.Reverse
                If (Fc.Extremity(espExtremityStart).y < 0) Then Fc.Reverse
                Fc.Layer = lyNew
                Call mSelection.Add(Fc)
            End If
        Next
    
        Call mSelection.AlignPlane(Document.Planes(strWorkPlane), espAlignPlaneFromGlobalXYZ)
        Document.ActivePlane = Document.Planes(strWorkPlane)
        
        For Each fcTemp In Document.FeatureChains
            If (fcTemp.Layer.Name = strTempLayer) Then
                fcTemp.Layer = lyTemp
            End If
        Next
        
        For Each ly In Document.Layers
            If (ly.Name = strTempLayer) Then
                Call Document.Layers.Remove(strTempLayer)
            End If
        Next
        
        lyNew.Visible = True
        'Document.Refresh
        
        'Tolerance back to original value
        Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bOriTolerance
        
        Exit Function
    
    End If

    '
    ' loop through all of the newly created features and reverse the closed ones
    '
    For i = LBound(GraphicObj) To UBound(GraphicObj)
        If GraphicObj(i).GraphicObjectType = espFeatureChain Then
            Set Fc = GraphicObj(i)
            'If FC.IsClosed Then FC.Reverse
            If (Fc.Extremity(espExtremityStart).y < 0) Then Fc.Reverse
        
            mSelection.DeleteAll
            'Call Document.GraphicsCollection.Remove(.GraphicsCollectionIndex)
            With mSelection
                .RemoveAll
                .AddCopiesToSelectionSet = False
            End With
            Call mSelection.Add(Fc)
        End If
    Next
    
    Call mSelection.AlignPlane(Document.Planes(strWorkPlane), espAlignPlaneFromGlobalXYZ)
    Document.ActivePlane = Document.Planes(strWorkPlane)
    'Document.Refresh
    
    Fc.Grouped = True
    'Document.Refresh
    
''''''''''''''''''''''''''''''''''''''''''''''''''
'' make Toolpath (3_ROUGH_ENDMILL)

    Dim FileName As String
    FileName = "3_ROUGH_ENDMILL.prc"
'    Dim feature As Esprit.FeatureChain
'    Set feature = Document.FeatureChains.Item(1)
    Dim M_TechnologyUtility As EspritTechnology.TechnologyUtility
    Set M_TechnologyUtility = Document.TechnologyUtility
    Dim tech() As EspritTechnology.Technology
    tech = M_TechnologyUtility.OpenProcess(EspritUserFolder & FileName)
    
    Dim Op As Esprit.Operation
    Set Op = Document.Operations.Add(tech(0), Fc)
    Op.Name = "3-" + strOperationOrder + ". ROUGH_ENDMILL_" & strWorkPlane

    Fc.Layer = lyOri
    Op.Layer = lyOri
    Document.ActiveLayer = lyOri
    Document.ActiveLayer.Visible = True
    Document.ActivePlane = Document.Planes("0DEG")
    
    For Each ly In Document.Layers
        If (ly.Name = strTempLayer) Then
            Call Document.Layers.Remove(strTempLayer)
        End If
    Next

    strTempLayer = "[" + strWorkPlane + "]" + Document.ActiveLayer.Name
    For Each ly In Document.Layers
        If (ly.Name = strTempLayer) Then
            Call Document.Layers.Remove(strTempLayer)
        End If
    Next
    Set lyTemp = Document.Layers.Add(strTempLayer)
    Fc.Layer = lyTemp
    Op.Layer = lyTemp
    lyTemp.Visible = True

    Document.Refresh
    
    'Tolerance back to original value
    Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bOriTolerance

End Function

'Sub Step3_3(Optional strLayerName As String = "ROUGH ENDMILL R6.0")
Sub Step3_3()
'연결된 선을 [자동연결]로 연결 피처를 생성
On Error Resume Next

    Dim graphicObject As Esprit.graphicObject
    Dim nCount As Integer
    Dim strLayerName As String
    strLayerName = "ROUGH ENDMILL R6.0"
    nCount = Document.Group.Count
    nCount = Document.SelectionSets.Count
   
'2. Select ;
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item(strLayerName)
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add(strLayerName)
    End With
    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
    End With
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = strLayerName And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType = espArc Or .GraphicObjectType = espPoint Or .GraphicObjectType = espSegment)) Then
            If (.Key > 0) Then
                .Grouped = True
                Call mSelection.Add(graphicObject)
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
    
    nCount = Document.Group.Count

    '
    ' create the feature chains
    '
    Dim GraphicObj() As Esprit.graphicObject
    If mSelection.Count > 0 Then
        GraphicObj = Document.FeatureRecognition.CreateAutoChains(mSelection)
    End If
    
    mSelection.DeleteAll
    Document.Refresh
    '
    ' loop through all of the newly created features and reverse the closed ones
    '
    For i = LBound(GraphicObj) To UBound(GraphicObj)
        If GraphicObj(i).GraphicObjectType = espFeatureChain Then
            Dim Fc As Esprit.FeatureChain
            Set Fc = GraphicObj(i)
            'If FC.IsClosed Then FC.Reverse
            If (Fc.Extremity(espExtremityStart).y < 0) Then
                Fc.Reverse
            End If
        End If
    Next

    Fc.Grouped = True
    Document.Refresh
    
End Sub
'Sub GetPartProfile(nActiveLayerIndex As Integer, strNamePlane As String, Optional ByVal strLayerName As String = "STL")
Function GetPartProfile(lyRef As Esprit.Layer, strNamePlane As String, Optional ByVal strLayerName As String = "STL", Optional ByVal bTolerance As Double = 0.1) As Double
    On Error Resume Next
    On Error GoTo 0
        
    Dim layerObject As Esprit.Layer
    Dim graphicObject As Esprit.graphicObject
    Dim stlObject As Esprit.STL_Model
    Dim solidObject As Esprit.Solid
    'Select Group: STL
    For Each layerObject In Esprit.Document.Layers
        layerObject.Visible = False
    Next
    Document.Refresh
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = strLayerName And (.GraphicObjectType = espSTL_Model)) Then
            If (.Key > 0) Then
                Set stlObject = graphicObject
                .Grouped = True
'                .Layer.Visible = True
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
        
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("subGetPartProfile")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("subGetPartProfile")
    End With

    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(stlObject)  'Add the stl_model to the selection object
    End With
    
    If mSelection.Count = 0 Then
        GetPartProfile = -1 'Error cannot find a STL model
        Exit Function
    End If
    
    Dim mRefGraphicObject() As Esprit.graphicObject
    Dim returnedFC As Esprit.FeatureChain
    
    Dim startPoint As Esprit.Point
    Dim midPoint As Esprit.Point
    Dim endPoint As Esprit.Point
    
    Dim lyOri As Esprit.Layer
    Set lyOri = Document.ActiveLayer
    
    'Tolerance set to 0.1 (default 공차)
    Dim bOriTolerance As Double
    bOriTolerance = Application.Configuration.ConfigurationFeatureRecognition.Tolerance
    If Document.SystemUnit = espInch Then
        Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bTolerance * 3.9
    Else
        Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bTolerance
    End If
    
    With Document
        .ActiveLayer = lyRef 'Set Turning Profile Layer to Active
        mRefGraphicObject = .FeatureRecognition.CreatePartProfileShadow(mSelection, Document.Planes.Item(strNamePlane), espSegmentsArcs)
    End With
    Document.ActiveLayer = lyOri
    Document.Refresh
    'mSelection = Nothing
    
    'Tolerance back to original value
    Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bOriTolerance
    
    GetPartProfile = 0
    
End Function

Function GetSTLYEnd(nDirectionCode As Integer, strNamePlane As String, Optional ByVal strLayerName As String = "STL") As Double
'Get connection width from Layer 10. Cut Off
        
'nDirectionCode = 0 : Bottom / 1 : Top / else : Error
    On Error Resume Next

    GetSTLYEnd = 0

    On Error GoTo 0
    
    
    'DirectionCode Check
    If Not (nDirectionCode = dBottom Or nDirectionCode = dTop) Then
        GoTo GetSTLYEnd
    End If
    
    
    Dim graphicObject As Esprit.graphicObject
    Dim stlObject As Esprit.STL_Model
    Dim dReturn As Double
    
    Dim lyrObjectInitial As Esprit.Layer
    Dim lyrObject As Esprit.Layer
    Set lyrObjectInitial = Document.ActiveLayer
    
    
    'For Test
    'Dim strSTLFilePath As String
    'strSTLFilePath = ".\Avaneer_Nickowski_Zmr 3.5_0_KL.stl"
    'strSTLFilePath = ".\Avaneer_Tech_01_Williams_Hiossen reg_0.stl"
    'Call GetSTL(strSTLFilePath)
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = strLayerName And (.GraphicObjectType = espSTL_Model)) Then
            If (.Key > 0) Then
                Set stlObject = graphicObject
            End If
        End If
        End With
    Next
    
    If stlObject Is Nothing Then GetSTLYEnd = -999 'check if the object is valid
    
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("Temp")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("Temp")
    End With

    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(stlObject)  'Add the stl_model to the selection object
    End With
    
    
    Dim mRefGraphicObject() As Esprit.graphicObject
    Dim returnedFC As Esprit.FeatureChain
    
    Dim startPoint As Esprit.Point
    Dim midPoint As Esprit.Point
    Dim endPoint As Esprit.Point
    
    With Document
        mRefGraphicObject = .FeatureRecognition.CreatePartProfileShadow(mSelection, Document.Planes.Item(strNamePlane), espFeatureChains)

        For Each graphicObject In Esprit.Document.GraphicsCollection
            With graphicObject
            If (.Layer.Visible And .Layer.Name = strLayerName And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType <> espWorkCoordinate)) Then
                If (.Key > 0) Then
                    .Grouped = False
                End If
            End If
            End With
        Next
        
        mRefGraphicObject(0).Grouped = True
        'plottedObjects = .GraphicsCollection.AddElements(comCurves)
        Set returnedFC = .FeatureChains.Item(.FeatureChains.Count)
        
        'Set startPoint = returnedFC.Extremity(espExtremityStart)
        'Set midPoint = returnedFC.Extremity(espExtremityMiddle)
        'Set endPoint = returnedFC.Extremity(espExtremityEnd)
    End With
    
    'Call mSelection.Translate(10, 0, 0)
    Dim graphicTemp As Esprit.graphicObject
    Dim lnLine As Esprit.Line
    Dim sgSegment As Esprit.Segment
    Dim dTopEnd As Double
    Dim dBottomEnd As Double
    Dim i As Integer
    
    dTopEnd = 0
    dBottomEnd = 0
    'For Each lnLine In Esprit.Document.Lines
    For i = 1 To returnedFC.Count
        
         If returnedFC.Item(i).TypeName = "Line" Then
            lnLine = returnedFC.Item(i)
            dBottomEnd = lnLine.y
            dTopEnd = lnLine.y
         
            'BottomEnd
            If lnLine.y > dBottomEnd Then
                dBottomEnd = lnLine.y
            End If
            'TopEnd
            If lnLine.y < dTopEnd Then
                dTopEnd = lnLine.y
            End If
         
         ElseIf returnedFC.Item(i).TypeName = "Segment" Then
            Set sgSegment = returnedFC.Item(i)
            'dBottom
            If minValue(sgSegment.YEnd, sgSegment.YStart) < dBottomEnd Then
                dBottomEnd = minValue(sgSegment.YEnd, sgSegment.YStart)
            End If
            'dTop
            If maxValue(sgSegment.YEnd, sgSegment.YStart) > dTopEnd Then
                dTopEnd = maxValue(sgSegment.YEnd, sgSegment.YStart)
            End If
        
        End If
        
    Next
    
    Dim SS As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set SS = .Item("tmpPartProfile")
        On Error GoTo 0
        If SS Is Nothing Then Set SS = .Add("tmpPartProfile")
    End With

    With SS
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(mRefGraphicObject)  'Add the stl_model to the selection object
    End With
    
    Call SS.Delete
    
    'Get Y Bottom End
    'dReturn = returnedFC.BoundingBoxLength + startPoint.X
    'nDirectionCode = 0 : Left / 1 : Right / else : Error
    If nDirectionCode = dBottom Then
        dReturn = dBottomEnd
    ElseIf nDirectionCode = dTop Then
        dReturn = dTopEnd
    Else
        dReturn = -999 'error
    End If
    'GetCutOffXRightEnd = dCutOffX(3)
    Document.ActiveLayer = lyrObjectInitial
    Document.Refresh
    
GetSTLYEnd:
    GetSTLYEnd = dReturn
End Function

Function GetLeftSegmentFromConnectionCuffOff(Optional ByVal nLayerNumber As Integer = 10) As Esprit.Segment

    Dim sgTemp As Esprit.Segment
    Dim sgReturn As Esprit.Segment
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Number = nLayerNumber And .TypeName = "Segment") Then
            If (.Key > 0) Then
                Set sgTemp = graphicObject
                If sgReturn Is Nothing Then
                    Set sgReturn = sgTemp
                Else
                    If minValue(sgTemp.XEnd, sgTemp.XStart) > 0 And (minValue(sgTemp.XEnd, sgTemp.XStart) < minValue(sgReturn.XEnd, sgReturn.XStart)) Then
                    Set sgReturn = sgTemp
                    End If
                End If
            End If
        End If
        End With
    Next

    Set GetLeftSegmentFromConnectionCuffOff = sgReturn
End Function

Sub checkInCircle()


    Dim pIntersect As Esprit.Point
    Dim go2 As Esprit.graphicObject
    
    Dim sg1 As Esprit.Segment
    Dim arc1 As Esprit.Arc
    
    Dim c1 As Esprit.Circle
    
    For Each cCircle In Esprit.Document.Circles
        With cCircle
        If (.Layer.Name = "기본값" And .TypeName = "Circle") Then
            If (.Key > 0) Then
                Set c1 = cCircle
                Exit For
            End If
        End If
        End With
    Next
    If Not (c1 Is Nothing) Then c1.Grouped = True
    
    
    For Each goFor In Esprit.Document.GraphicsCollection
        With goFor
        If (.Layer.Name = "STL" And (.TypeName = "Segment" Or .TypeName = "Arc")) Then
            Set go2 = goFor
            Set pIntersect = IntersectCircleAndArcsSegments(c1, go2)
        End If
        End With
    Next

Dim i As Integer
i = 1

End Sub


Public Function EspritUserFolder() As String
    ' the contents of this folder are saved with the .esp file
    EspritUserFolder = Application.TempDir & "USER\"
'    EspritUserFolder = Application.Configuration.GetFileDirectory(espFileTypeTemplate) & "\TechPrcs\"

End Function

Public Function GetTechnologyUtility() As EspritTechnology.TechnologyUtility
    ' this function just serves to typecast Document.TechnologyUtility
    ' from a generic VB Object to an EspritTechnology.TechnologyUtility
    Set GetTechnologyUtility = Document.TechnologyUtility
End Function

Sub DeleteSmallPointChainFeature(strLayerName As String, nLen As Double)

    Dim lyTemp As Esprit.Layer
    Dim fcTemp As Esprit.FeatureChain
    Dim strTempLayer As String
    Dim ly As Esprit.Layer

    For Each ly In Document.Layers
        If (ly.Name = strLayerName) Then
            Set lyTemp = ly
        End If
    Next
    
    Document.ActiveLayer = lyTemp
    Document.ActiveLayer.Visible = True
    
    'Count FC & delete less than the length
    Dim nCnt As Integer
    nCnt = 0
    For Each fcCnt In Document.FeatureChains
        If (fcCnt.Layer.Name = lyTemp.Name) Then
            Set fcTemp = fcCnt
            If (fcCnt.Length > nLen) Then
                nCnt = nCnt + 1
            Else
                Document.FeatureChains.Remove (fcTemp.Key)
            End If
            
        End If
    Next
    'Document.Refresh
    
End Sub



Function PublicfindTextFeatureChain() As FeatureChain
    Dim chnText As FeatureChain
    Dim rtnText As FeatureChain
    
    PublicsetLayersForText
    Set lyOri = Document.ActiveLayer

    For Each chnText In Esprit.Document.FeatureChains
        If (chnText.Layer.Name = lyOri.Name) Then
            Set rtnText = chnText
        End If
    Next
    
    Set PublicfindTextFeatureChain = rtnText
End Function

Function PublicfindTextSelectionSet() As SelectionSet
    Dim strmSelectionIndex As String
    strmSelectionIndex = strWorkPlane + "PGMText"
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
    
    PublicsetLayersForText
    Set lyOri = Document.ActiveLayer

    For Each goRef In Esprit.Document.GraphicsCollection
        With goRef
        If (.Layer.Name = lyOri.Name And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType = espFeatureChain)) Then
            If (.Key > 0) Then
                .Grouped = True
                Call mSelection.Add(goRef)
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
    
    Set PublicfindTextSelectionSet = mSelection
    
End Function

Sub PublicsetLayersForText()

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

End Sub


Sub createProjectionFinishing()
    Dim strWorkPlaneName As String
    strWorkPlaneName = Document.ActivePlane.Name
    Dim strBaseLayerName As String
    strBaseLayerName = Document.ActiveLayer.Name
    
    Dim Fc As FeatureChain
    Set Fc = PublicfindTextFeatureChain
    
    Dim mSelection As Esprit.SelectionSet
    Set mSelection = PublicfindTextSelectionSet
    
    If mSelection.Count = 0 Then
        MsgBox ("Cannot find the Text. Please generate it first.")
        Exit Sub
    End If

''''''''''''''''''''''''''''''''''''''''''''''''''
'' make Toolpath (Projection Finishing)
    
    Dim FileName As String
    FileName = "TEXT_Projection_FInishing.prc"
    
    Dim M_TechnologyUtility As EspritTechnology.TechnologyUtility
    Set M_TechnologyUtility = Document.TechnologyUtility
    Dim tech() As EspritTechnology.Technology
    tech = M_TechnologyUtility.OpenProcess(EspritUserFolder & FileName)
    'Set Tech = M_TechnologyUtility.CreateTechnology(espTechLatheMill3DProject, espInch)
    
    
    Dim Op As Esprit.Operation
    Set Op = Document.Operations.Add(tech(0), Fc)

    Op.Name = "9. " & strWorkPlane & " TEXT(PGMNo):" & GetProgramNumber() & " CROSS BALL ENDMILL R0.2"

'1)  Layer를 strBaseLayerName: TEXT 으로 선택
    Dim lyOri As Esprit.Layer
    Dim ly As Esprit.Layer
    Dim strOriginalLayer As String
    strOriginalLayer = strBaseLayerName
    
    For Each ly In Document.Layers
        If (ly.Name = strOriginalLayer) Then
            Set lyOri = ly
        End If
    Next
    
    Document.ActiveLayer = lyOri
    Document.ActiveLayer.Visible = True

    Fc.Layer = lyOri
    Op.Layer = lyOri
    Document.ActiveLayer = lyOri
    Document.ActiveLayer.Visible = True

    Document.Refresh
End Sub

Sub movePlaneNext(param As udfPlaneDegree)
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
        nDegree = (CInt(strDegree) + param) Mod 360
        strDegree = CStr(nDegree) + "DEG"
        
        Document.ActivePlane = Document.Planes(strDegree)
    End If

    Document.Refresh
End Sub
Sub movePlaneBefore(param As udfPlaneDegree)
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
        nDegree = (CInt(strDegree) + 360 - param) Mod 360
        strDegree = CStr(nDegree) + "DEG"
        
        Document.ActivePlane = Document.Planes(strDegree)
    End If

    Document.Refresh
End Sub
