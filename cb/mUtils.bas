Attribute VB_Name = "mUtils"
'mUtils
Option Private Module


Function PI() As Double

    PI = 4 * Atn(1)

End Function

Function getTheOriginAxis() As Esprit.Line
    Dim L As Esprit.Line
    With Document.ActivePlane
        Set L = Document.GetLine(Document.GetPoint(.x, .y, .Z), .Wx, .Wy, .Wz) 'Active plane's W vector
    End With
    
    Set getTheOriginAxis = L
End Function

Public Sub ClickToContinue()

    ' temporarily turns on grid mode to let the user left click in graphics to continue

    Dim PreviousGridState As Boolean

    PreviousGridState = Application.Configuration.EnableGridMode

    Application.Configuration.EnableGridMode = True

    On Error Resume Next

    Call Document.GetAnyElement("Press Left Mouse Button in Graphics to Continue", espPoint)

    Application.Configuration.EnableGridMode = PreviousGridState

End Sub

 

Public Function Distance2ComPoints(Pt1 As EspritGeometry.ComPoint, Pt2 As EspritGeometry.ComPoint) As Double

    Distance2ComPoints = Sqr((Pt2.x - Pt1.x) ^ 2 + (Pt2.y - Pt1.y) ^ 2 + (Pt2.Z - Pt1.Z) ^ 2)

End Function

 

Public Sub SlugRemoval()

    '

    ' take the one and only stock and split it

    '

    Dim Stock As EspritSimulation.SimuStock

    Set Stock = MySimu.Stocks.Item(1)

    Call Stock.Split

    '

    ' find the bounding boxes of the stocks that result

    '

    Dim P(1 To 2) As EspritGeometry.ComPoint

    Set P(1) = New EspritGeometry.ComPoint

    Set P(2) = New EspritGeometry.ComPoint

    Dim NumberOfStocks As Long, i As Long, Distance() As Double, LargestStock As Long

    ' only 2 stocks should result, but set up for more

    NumberOfStocks = MySimu.Stocks.Count

    ReDim Distance(0 To NumberOfStocks) As Double

    For i = 1 To NumberOfStocks

        Set Stock = MySimu.Stocks.Item(i)

        Call Stock.GetBoundingBox(P(1), P(2))

        Distance(i) = Distance2ComPoints(P(1), P(2))

        ' use distance(0) to track largest stock

        If Distance(i) > Distance(0) Then

            Distance(0) = Distance(i)

            LargestStock = i

        End If

    Next

    '

    ' loop back through the and remove all but the largest stock

    '

    For i = NumberOfStocks To 1 Step -1

        If i <> LargestStock Then Call MySimu.Stocks.Remove(i)

    Next

    Call Document.Windows.ActiveWindow.Refresh

End Sub

 

Public Function Radians(Degrees As Double) As Double

    Radians = PI * Degrees / 180

End Function

 

Public Function FindToolByNumber(ToolNumber As Long, Optional ExpandSearch As Boolean = True) As EspritTechnology.Technology

    Dim T As EspritTechnology.Technology, TN As Long

    For Each T In Document.Tools

        TN = 0

        On Error Resume Next

        TN = T.Item("ToolNumber").Value

        On Error GoTo 0

        If TN = ToolNumber Then

            Set FindToolByNumber = T

            Exit Function

        End If

    Next

    '

    ' if it is more than 2 digits, only try first digits

    '

    If ExpandSearch And (ToolNumber > 100) Then

        Set FindToolByNumber = FindToolByNumber(Int(ToolNumber / 100))

    End If

End Function

 
'
'Private Function ExportAllToolsAndReturnFirst() As EspritSimulation.ISimuTool
'
'    If MySimu Is Nothing Then Exit Function
'
'    If MySimu.Tools.Count = 0 Then
'
'        Call Document.Simulation.ExportAllToolsToSimulation(MySimu)
'
'    End If
'
'    If MySimu.Tools.Count > 0 Then
'
'        Set ExportAllToolsAndReturnFirst = MySimu.Tools.Item(1)
'
'    End If
'
'End Function



' ------------------------------------------------------------------------
' Copyright (C) 1998-2005 DP Technology Corp.
'
' This sample code file is provided for illustration purposes. You have a
' royalty-free right to use, modify, reproduce and distribute the sample
' code files (and/or any modified versions) in any way you find useful,
' provided that you agree that DP Technology has no warranty obligations
' or liability whatsoever for any of the sample code files, the result
' from your use of them or for any modifications you may make. This notice
' must be included on any reproduced or distributed file or modified file.
' ------------------------------------------------------------------------
'modComGeoUtilities
Function PointToComPoint(PointObj As Esprit.Point) As EspritGeometry.ComPoint
    If PointObj Is Nothing Then Exit Function
    ' returns a ComPoint object corresponding to a given Esprit.Point
    Set PointToComPoint = New EspritGeometry.ComPoint
    With PointObj
        Call PointToComPoint.SetXyz(.x, .y, .Z)
    End With
End Function
Function LineToComLine(LineObj As Esprit.Line) As EspritGeometry.ComLine
    If LineObj Is Nothing Then Exit Function
    ' returns a ComLine object corresponding to a given Esprit.Line
    Set LineToComLine = New EspritGeometry.ComLine
    With LineObj
        Call LineToComLine.Point.SetXyz(.x, .y, .Z)
        Call LineToComLine.Vector.SetXyz(.Ux, .Uy, .Uz)
    End With
End Function

Function CircleToComCircle(CircleObj As Esprit.Circle) As EspritGeometry.ComCircle
    If CircleObj Is Nothing Then Exit Function
    ' returns a ComCircle object corresponding to a given Esprit.Circle
    Set CircleToComCircle = New EspritGeometry.ComCircle
    With CircleObj
        Call CircleToComCircle.CenterPoint.SetXyz(.x, .y, .Z)
        CircleToComCircle.Radius = .Radius
        Call CircleToComCircle.U.SetXyz(.Ux, .Uy, .Uz)
        Call CircleToComCircle.V.SetXyz(.Vx, .Vy, .Vz)
    End With
End Function

Function SegmentToComSegment(SegmentObj As Esprit.Segment) As EspritGeometry.ComSegment
    If SegmentObj Is Nothing Then Exit Function
    ' returns a ComSegment object corresponding to a given Esprit.Segment
    Set SegmentToComSegment = New EspritGeometry.ComSegment
    With SegmentObj
        Call SegmentToComSegment.startPoint.SetXyz(.XStart, .YStart, .ZStart)
        Call SegmentToComSegment.endPoint.SetXyz(.XEnd, .YEnd, .ZEnd)
    End With
End Function

Function ArcToComArc(ArcObj As Esprit.Arc) As EspritGeometry.ComArc
    If ArcObj Is Nothing Then Exit Function
    ' returns a ComArc object corresponding to a given Esprit.Arc
    Set ArcToComArc = New EspritGeometry.ComArc
    With ArcObj
        Call ArcToComArc.CenterPoint.SetXyz(.x, .y, .Z)
        ArcToComArc.Radius = .Radius
        ArcToComArc.StartAngle = .StartAngle
        ArcToComArc.EndAngle = .EndAngle
        Call ArcToComArc.U.SetXyz(.Ux, .Uy, .Uz)
        Call ArcToComArc.V.SetXyz(.Vx, .Vy, .Vz)
    End With
End Function

Function GraphicObjectToComGeoBase(GraphicObj As Esprit.graphicObject) As EspritGeometryBase.ComGeoBase
    If GraphicObj Is Nothing Then Exit Function
    ' returns a ComGeoBase object corresponding to a given Esprit.GraphicObject
    Select Case GraphicObj.GraphicObjectType
    Case espArc
        Set GraphicObjectToComGeoBase = ArcToComArc(GraphicObj)
    Case espCircle
        Set GraphicObjectToComGeoBase = CircleToComCircle(GraphicObj)
    Case espLine
        Set GraphicObjectToComGeoBase = LineToComLine(GraphicObj)
    Case espPoint
        Set GraphicObjectToComGeoBase = PointToComPoint(GraphicObj)
    Case espSegment
        Set GraphicObjectToComGeoBase = SegmentToComSegment(GraphicObj)
    End Select
End Function

Function ComPointToPoint(ComPointObj As EspritGeometry.ComPoint, Optional CreateVirtual As Boolean = False) As Esprit.Point
    If ComPointObj Is Nothing Then Exit Function
    ' creates and returns an Esprit.Point from a given ComPoint
    With ComPointObj
        If CreateVirtual Then
            Set ComPointToPoint = Document.GetPoint(.x, .y, .Z)
        Else
            Set ComPointToPoint = Document.Points.Add(.x, .y, .Z)
        End If
    End With
End Function

Function ComLineToLine(ComLineObj As EspritGeometry.ComLine, Optional CreateVirtual As Boolean = False) As Esprit.Line
    If ComLineObj Is Nothing Then Exit Function
    ' creates and returns an Esprit.Line from a given ComLine
    With ComLineObj
        Dim TempP As Esprit.Point
        With .Point
            Set TempP = Document.GetPoint(.x, .y, .Z)
        End With
        With .Vector
            If CreateVirtual Then
                Set ComLineToLine = Document.GetLine(TempP, .x, .y, .Z)
            Else
                Set ComLineToLine = Document.Lines.Add(TempP, .x, .y, .Z)
            End If
        End With
    End With
End Function

Function ComCircleToCircle(ComCircleObj As EspritGeometry.ComCircle, Optional CreateVirtual As Boolean = False) As Esprit.Circle
    If ComCircleObj Is Nothing Then Exit Function
    ' creates and returns an Esprit.Circle from a given ComCircle
    With ComCircleObj
        Dim TempCP As Esprit.Point
        With .CenterPoint
            Set TempCP = Document.GetPoint(.x, .y, .Z)
        End With
        If CreateVirtual Then
            Set ComCircleToCircle = Document.GetCircle(TempCP, .Radius)
        Else
            Set ComCircleToCircle = Document.Circles.Add(TempCP, .Radius)
        End If
        With .U
            ComCircleToCircle.Ux = .x
            ComCircleToCircle.Uy = .y
            ComCircleToCircle.Uz = .Z
        End With
        With .V
            ComCircleToCircle.Vx = .x
            ComCircleToCircle.Vy = .y
            ComCircleToCircle.Vz = .Z
        End With
    End With
End Function

Function ComSegmentToSegment(ComSegmentObj As EspritGeometry.ComSegment, Optional CreateVirtual As Boolean = False) As Esprit.Segment
    If ComSegmentObj Is Nothing Then Exit Function
    ' creates and returns an Esprit.Segment from a given ComSegment
    With ComSegmentObj
        Dim TempP(1 To 2) As Esprit.Point
        With .startPoint
            Set TempP(1) = Document.GetPoint(.x, .y, .Z)
        End With
        With .endPoint
            Set TempP(2) = Document.GetPoint(.x, .y, .Z)
        End With
        If CreateVirtual Then
            Set ComSegmentToSegment = Document.GetSegment(TempP(1), TempP(2))
        Else
            Set ComSegmentToSegment = Document.Segments.Add(TempP(1), TempP(2))
        End If
    End With
End Function

Function ComArcToArc(ComArcObj As EspritGeometry.ComArc, Optional CreateVirtual As Boolean = False) As Esprit.Arc
    If ComArcObj Is Nothing Then Exit Function
    ' creates and returns an Esprit.Arc from a given ComArc
    With ComArcObj
        Dim TempCP As Esprit.Point
        With .CenterPoint
            Set TempCP = Document.GetPoint(.x, .y, .Z)
        End With
        If CreateVirtual Then
            Set ComArcToArc = Document.GetArc(TempCP, .Radius, .StartAngle, .EndAngle)
        Else
            Set ComArcToArc = Document.Arcs.Add(TempCP, .Radius, .StartAngle, .EndAngle)
        End If
        With .U
            ComArcToArc.Ux = .x
            ComArcToArc.Uy = .y
            ComArcToArc.Uz = .Z
        End With
        With .V
            ComArcToArc.Vx = .x
            ComArcToArc.Vy = .y
            ComArcToArc.Vz = .Z
        End With
    End With
End Function

Function GetComPoint(x As Double, y As Double, Z As Double) As EspritGeometry.ComPoint
    ' creates and returns a ComPoint having the given coordinates
    Set GetComPoint = New EspritGeometry.ComPoint
    Call GetComPoint.SetXyz(x, y, Z)
End Function

Function GetComVector(x As Double, y As Double, Z As Double) As EspritGeometry.ComVector
    ' creates and returns a ComVector having the given components
    Set GetComVector = New EspritGeometry.ComVector
    Call GetComVector.SetXyz(x, y, Z)
End Function

Function GetComMatrix(U As EspritGeometry.ComVector, V As EspritGeometry.ComVector, W As EspritGeometry.ComVector) As EspritGeometry.ComMatrix
    ' creates and returns a ComMatrix having the given vectors
    Set GetComMatrix = New EspritGeometry.ComMatrix
    With GetComMatrix
        .U = U
        .V = V
        .W = W
    End With
End Function



Function maxValue(A As Double, B As Double) As Double
    If A >= B Then
        maxValue = A
    Else
        maxValue = B
    End If
End Function



Function minValue(A As Double, B As Double) As Double
    If A <= B Then
        minValue = A
    Else
        minValue = B
    End If
End Function


Function GetGeoUtility() As EspritGeometryRoutines.GeoUtility
    Static MyGeoUtility As EspritGeometryRoutines.GeoUtility
    If MyGeoUtility Is Nothing Then
        Set MyGeoUtility = New EspritGeometryRoutines.GeoUtility
    End If
    Set GetGeoUtility = MyGeoUtility
End Function


Function IntersectSegments(ByRef Segment1 As Esprit.Segment, _
            ByRef Segment2 As Esprit.Segment) As Esprit.Point
    Dim ComSeg(1) As EspritGeometry.ComSegment
    Dim IntPt() As EspritGeometry.ComPoint
    Set ComSeg(0) = SegmentToComSegment(Segment1)
    Set ComSeg(1) = SegmentToComSegment(Segment2)
    IntPt = GetGeoUtility.Intersect(ComSeg(0).Unbound, ComSeg(1).Unbound)
    Set IntersectSegments = ComPointToPoint(IntPt(0), True)
End Function



Function IntersectSegment1Arc1(ByRef pSegment1 As Esprit.Segment, _
            ByRef pArc2 As Esprit.Arc) As Esprit.Point
    Dim ComSeg1 As EspritGeometry.ComSegment
    Dim ComArc2 As EspritGeometry.ComArc
    Dim IntPt() As EspritGeometry.ComPoint
    Set ComSeg1 = SegmentToComSegment(pSegment1)
    Set ComArc2 = ArcToComArc(pArc2)
    IntPt = GetGeoUtility.Intersect(ComSeg1.Unbound, ComArc2.Unbound)
    Set IntersectSegment1Arc1 = ComPointToPoint(IntPt(0), True)
End Function


Function IntersectCircleAndArcsSegments(ByRef cBound As Esprit.Circle, _
            ByRef go2 As Esprit.graphicObject) As Esprit.Point
            
    On Error GoTo ErrorHandler
    Dim ComCircleBound As EspritGeometry.ComCircle
    
    Dim ComSeg(1) As EspritGeometry.ComSegment
    Dim ComArc(1) As EspritGeometry.ComArc
    Dim sgTmp As Esprit.Segment
    Dim arcTmp As Esprit.Arc
    Dim IntPt() As EspritGeometry.ComPoint
    
    Set ComCircleBound = CircleToComCircle(cBound)
    
    If (go2.GraphicObjectType = espSegment) Then
        Set sgTmp = go2
        Set ComSeg(1) = SegmentToComSegment(sgTmp)
        IntPt = GetGeoUtility.Intersect(ComCircleBound, ComSeg(1).Unbound)
        
    Else
        If (go2.GraphicObjectType = espArc) Then
            Set arcTmp = go2
            Set ComArc(1) = ArcToComArc(arcTmp)
            IntPt = GetGeoUtility.Intersect(ComCircleBound, ComArc(1).Unbound)
        Else
            'Error(not supported GO)
            Return
        End If
    End If
    
    Dim i As Integer
    Dim ptTmp As Esprit.Point
    i = 0
    'Boundary: in circle radius.
    For i = 0 To UBound(IntPt) Step 1
        Set ptTmp = ComPointToPoint(IntPt(i))
        If (IntPt(i).x > cBound.Radius * (-1) And IntPt(i).x < cBound.Radius _
            And IntPt(i).y > cBound.Radius * (-1) And IntPt(i).y < cBound.Radius) Then
            If (go2.GraphicObjectType = espSegment) Then
                'For Debug
                'sgTmp.Grouped = True
                'ComPointToPoint(IntPt(i)).Grouped = True
                If (IsPointOnSegment(sgTmp, ptTmp)) Then
                    Set IntersectCircleAndArcsSegments = ComPointToPoint(IntPt(i), True)
                End If
            ElseIf (go2.GraphicObjectType = espArc) Then
                If IsPointOnArc(arcTmp, ptTmp) Then
                    Set IntersectCircleAndArcsSegments = ComPointToPoint(IntPt(i), True)
                End If
            End If
        End If
        Call Document.GraphicsCollection.Remove(ptTmp.GraphicsCollectionIndex)
    Next
    
    Exit Function
ErrorHandler:
    'IntersectCircleAndArcsSegments = Nothing
    If Not (ptTmp Is Nothing) Then
        Call Document.GraphicsCollection.Remove(ptTmp.GraphicsCollectionIndex)
    End If
    Exit Function
    
End Function

Function IntersectArcsSegments(ByRef go1 As Esprit.graphicObject, _
            ByRef go2 As Esprit.graphicObject) As Esprit.Point
    
    On Error GoTo ErrorHandler
    
    Dim ComSeg(1) As EspritGeometry.ComSegment
    Dim ComArc(1) As EspritGeometry.ComArc
    Dim sgTmp As Esprit.Segment
    Dim arcTmp As Esprit.Arc
    Dim IntPt() As EspritGeometry.ComPoint
    
    If (go1.GraphicObjectType = espSegment) Then
        Set sgTmp = go1
        Set ComSeg(0) = SegmentToComSegment(sgTmp)
    Else
        If (go1.GraphicObjectType = espArc) Then
            Set arcTmp = go1
            Set ComArc(0) = ArcToComArc(arcTmp)
        End If
    End If
    
    If (go2.GraphicObjectType = espSegment) Then
        Set sgTmp = go2
        Set ComSeg(1) = SegmentToComSegment(sgTmp)
        If Not (ComSeg(0) Is Nothing) Then
            IntPt = GetGeoUtility.Intersect(ComSeg(0).Unbound, ComSeg(1).Unbound)
        ElseIf Not (ComArc(0) Is Nothing) Then
                IntPt = GetGeoUtility.Intersect(ComArc(0).Unbound, ComSeg(1).Unbound)
        Else
            'Error(not supported GO)
            Return
        End If
    Else
        If (go2.GraphicObjectType = espArc) Then
            Set arcTmp = go2
            Set ComArc(1) = ArcToComArc(arcTmp)
            If Not (ComSeg(0) Is Nothing) Then
                IntPt = GetGeoUtility.Intersect(ComSeg(0).Unbound, ComArc(1).Unbound)
            ElseIf Not (ComArc(0) Is Nothing) Then
                IntPt = GetGeoUtility.Intersect(ComArc(0).Unbound, ComArc(1).Unbound)
            Else
                'Error(not supported GO)
                Return
            End If
        End If
    End If

'    Dim i As Integer
'    i = 0
'    '나중에 Boundary 값을 정해야 할 것.
'    For i = 0 To UBound(IntPt) Step 1
'        If (IntPt(i).X > -100 And IntPt(i).X < 100 _
'            And IntPt(i).Y > -100 And IntPt(i).Y < 100) Then
'            Set IntersectArcsSegments = ComPointToPoint(IntPt(i), True)
'        End If
'    Next
    
    Dim i As Integer
    Dim ptTmp As Esprit.Point
    i = 0
    'Boundary: in circle radius.
    For i = 0 To UBound(IntPt) Step 1
        Set ptTmp = ComPointToPoint(IntPt(i))
        If (go2.GraphicObjectType = espSegment) Then
            'For Debug
            'sgTmp.Grouped = True
            'ComPointToPoint(IntPt(i)).Grouped = True
            If (IsPointOnSegment(go2, ptTmp)) Then
                If (go1.GraphicObjectType = espSegment) Then
                    If (IsPointOnSegment(go1, ptTmp)) Then
                        Set IntersectArcsSegments = ComPointToPoint(IntPt(i), True)
                    End If
                ElseIf (go1.GraphicObjectType = espArc) Then
                    If (IsPointOnArc(go1, ptTmp)) Then
                        Set IntersectArcsSegments = ComPointToPoint(IntPt(i), True)
                    End If
                End If
            End If
        ElseIf (go2.GraphicObjectType = espArc) Then
            If IsPointOnArc(go2, ptTmp) Then
                If (go1.GraphicObjectType = espSegment) Then
                    If (IsPointOnSegment(go1, ptTmp)) Then
                        Set IntersectArcsSegments = ComPointToPoint(IntPt(i), True)
                    End If
                ElseIf (go1.GraphicObjectType = espArc) Then
                    If (IsPointOnArc(go1, ptTmp)) Then
                        Set IntersectArcsSegments = ComPointToPoint(IntPt(i), True)
                    End If
                End If
            End If
        End If
        Call Document.GraphicsCollection.Remove(ptTmp.GraphicsCollectionIndex)
    Next
    Exit Function

ErrorHandler:
    'IntersectCircleAndArcsSegments = Nothing
    If Not (ptTmp Is Nothing) Then
        Call Document.GraphicsCollection.Remove(ptTmp.GraphicsCollectionIndex)
    End If
    Exit Function
    
    
End Function



Public Function IsPointOnArc(ByVal Arc As Esprit.Arc, ByVal Point As Esprit.Point, _
                             Optional ByVal nearestPointOn As Esprit.Point) As Boolean
    Dim returnComPoint As EspritGeometry.ComPoint
    Dim bTolerance As Double
    bTolerance = 0.0001
    
    'IsPointOnArc = ArcToComArc(Arc).IsPointOn(PointToComPoint(Point), SystemTolerance, SystemTolerance, returnComPoint)
    IsPointOnArc = ArcToComArc(Arc).IsPointOn(PointToComPoint(Point), bTolerance, bTolerance, returnComPoint)
    'Set nearestPointOn = ConvertToPoint(returnComPoint)
    'Set nearestPointOn = ComPointToPoint(returnComPoint)
    
    
End Function

Public Function IsPointOnSegment(ByVal Segment As Esprit.Segment, ByVal Point As Esprit.Point, _
                                 Optional ByVal nearestPointOn As Esprit.Point) As Boolean
    Dim returnComPoint As EspritGeometry.ComPoint
    Dim bTolerance As Double
    bTolerance = 0.0001

'    IsPointOnSegment = SegmentToComSegment(Segment).IsPointOn(PointToComPoint(Point), SystemTolerance, SystemTolerance, returnComPoint)
    IsPointOnSegment = SegmentToComSegment(Segment).IsPointOn(PointToComPoint(Point), bTolerance, bTolerance, returnComPoint)
    'Set nearestPointOn = ConvertToPoint(returnComPoint)
    'Set nearestPointOn = ComPointToPoint(returnComPoint)
End Function



Public Sub SaveProcessDemo()
    Dim FileName As String
    FileName = Application.TempDir & "SaveProcessDemo.prc"
    Call SaveProcessGroupedOperations(FileName)
End Sub

Public Sub SaveProcessGroupedOperations(ByVal FileName As String)
    Dim Technology() As EspritTechnology.Technology
    ReDim Technology(Document.Operations.Count)
    Dim Op As Esprit.Operation, processCount As Long
    For Each Op In Document.Operations
        If Op.Grouped Then
            Set Technology(processCount) = Op.Technology
            processCount = processCount + 1
        End If
    Next
    If processCount > 0 Then
        ReDim Preserve Technology(processCount - 1)
        Dim techUtil As EspritTechnology.TechnologyUtility
        Set techUtil = Document.TechnologyUtility
        Call techUtil.SaveProcess(Technology, FileName)
    Else
        Call MsgBox("Please group select operations before running this macro.")
    End If
End Sub

Public Sub OpenProcessDemo()
    Dim FileName As String
    FileName = Application.TempDir & "SaveProcessDemo.prc"
    If FileSystem.Dir(FileName) <> vbNullString Then
        Dim techUtil As EspritTechnology.TechnologyUtility
        Set techUtil = Document.TechnologyUtility
        Dim Technology() As EspritTechnology.Technology
        Technology = techUtil.OpenProcess(FileName)
        Dim i As Long
        For i = LBound(Technology) To UBound(Technology)
            Debug.Print i, TypeName(Technology(i))
        Next
    Else
        Call MsgBox("File does not exist: " & vbNewLine & _
                    FileName & vbNewLine & vbNewLine & _
                    "Please run the SaveProcessDemo macro first to create it.")
    End If
End Sub



Sub SortReservoirOperations()
On Error GoTo errHandler
Dim lErr As Long, sErrDesc As String, sSource As String
Dim Op As Esprit.Operation
Dim i As Integer
Dim J As Integer
Dim dPrimaryAngle As Double
Dim Ops() As Esprit.Operation
Dim Fc As Esprit.FeatureChain
Dim tech As EspritTechnology.Technology
Dim XCoordinate As Double
Dim ZCoordinate As Double
Dim esArc As Esprit.Arc
Dim esPoint As Esprit.Point
Dim rsOpData As ADODB.Recordset
Dim myFieldList As Variant
Dim iFirstMachiningOP As Integer

    Set rsOpData = New ADODB.Recordset
    With rsOpData
        '   Create columns
        .Fields.Append "OPNumber", adInteger
        .Fields.Append "AAngle", adDouble
        .Fields.Append "XPostion", adDouble
        .Fields.Append "ZPostion", adDouble
        '   Set up array with field names because it's much more convenient when adding data
        myFieldList = Array("OPNumber", "AAngle", "XPostion", "ZPostion")
        .Open

        For i = 1 To gobjEspDoc.Operations.Count
            Set Op = gobjEspDoc.Operations(i)
            Set tech = Op.Technology
            If tech.TechnologyType <> espTechMillCustom And _
            tech.TechnologyType <> espTechMillManual Then
                iFirstMachiningOP = i
                Exit For
            End If
        Next i

        For i = iFirstMachiningOP To gobjEspDoc.Operations.Count
            Set Op = gobjEspDoc.Operations(i)
            dPrimaryAngle = Degrees(Op.AnglePrimary)
            
'            If dPrimaryAngle > 0 Then
'                dPrimaryAngle = dPrimaryAngle - 360
'            End If
            If dPrimaryAngle < 0 Then
                dPrimaryAngle = dPrimaryAngle + 360
            End If
            
            Set tech = Op.Technology
            If tech.TechnologyType <> espTechMillCustom And _
            tech.TechnologyType <> espTechMillManual Then
                If Op.Feature.GraphicObjectType = espFeatureChain Then
                    Set Fc = Op.Feature
                    For J = 1 To Fc.Count
                        If Fc.Item(J).GraphicObjectType = espArc Then
                            Set esArc = Fc.Item(1)
                            XCoordinate = esArc.CenterPoint.x
                            Exit For
                        End If
                    Next J
                    ZCoordinate = Fc.MaxZ
                End If
            Else
                XCoordinate = 0
                ZCoordinate = 0
            End If
            .AddNew myFieldList, Array(i, Abs(dPrimaryAngle), XCoordinate, ZCoordinate)
            .Update
'            Debug.Print "XCoordinate " & XCoordinate
'            Debug.Print "ZCoordinate " & ZCoordinate
        Next i
    
        .Sort = "AAngle, XPostion desc, ZPostion desc "
        
        ReDim Ops(1 To gobjEspDoc.Operations.Count - (iFirstMachiningOP - 1)) As Esprit.Operation
        
        .MoveFirst
        For i = 1 To .RecordCount
'            Debug.Print .AbsolutePosition & " Cursor Postion"
'            Debug.Print "OpNum " & !OpNumber, "AAngle " & !AAngle, "X " & !XPostion, "Z " & !ZPostion
            J = !OpNumber
            Set Ops(i) = gobjEspDoc.Operations.Item(J)
           .Move i, 1
        Next i
    End With
        
    Call gobjEspDoc.Operations.Reorder(Ops, iFirstMachiningOP, True)
    DoEvents
    
rtnExit:
    On Error GoTo 0
    
    rsOpData.Close
    Set rsOpData = Nothing
    Set esArc = Nothing
    Set Fc = Nothing
    Erase Ops()
    
    On Error GoTo 0

    If lErr <> 0 Then
        Err.Raise lErr, , "SortReservoirOperations Error " & sErrDesc
    End If

    Exit Sub

errHandler:
    ' save our error
    lErr = Err.Number
    sErrDesc = Err.Description
    sSource = Err.Source

    Resume rtnExit
    Resume Next

End Sub


Sub SaveOperationOrder()
    ' saves the operation list order by tagging a custom property
    ' on each op with its index position in the operations collection
    Dim i As Long, Op As Esprit.Operation
    With Document.Operations
        For i = 1 To .Count
            Set Op = .Item(i)
            Call SetCustomLong(Op.CustomProperties, "SortOrder", i)
        Next
    End With
End Sub
Function GetCustomLong(ByRef CustomProps As EspritProperties.CustomProperties, _
                       PropertyName As String) As EspritProperties.CustomProperty
    With CustomProps
        On Error Resume Next  ' in case the property does not exist
        Set GetCustomLong = .Item(PropertyName)
    End With
End Function

Function SetCustomLong(ByRef CustomProps As EspritProperties.CustomProperties, _
                       PropertyName As String, NewValue As Long) As EspritProperties.CustomProperty
    With CustomProps
        On Error Resume Next
        Call .Remove(PropertyName) ' in case the property already exists
        Set SetCustomLong = .Add(PropertyName, PropertyName, espPropertyTypeLong, NewValue)
    End With
End Function

Sub DeleteDummyOperation()
    
    Dim gclist() As Long
    
    Dim i, d As Long, Op As Esprit.Operation
    d = 0
    ReDim gclist(0 To Document.Operations.Count)
    With Document.Operations
        For i = 1 To .Count
            Set Op = .Item(i)
            If InStr(1, Op.Name, "DUMMY") Then
                gclist(d) = Op.GraphicsCollectionIndex
                d = d + 1
            End If
        Next
    End With
    
    For x = 0 To (d - 1)
        Call Document.GraphicsCollection.Remove(gclist(x))
    Next

End Sub

Sub RestoreLastSavedOperationOrder()
    ' sorts the operations collection by comparing the custom property
    ' tags that were saved by the SaveOperationOrder macro above
    ' any new operations created since the last save are pushed to the bottom
    Dim i As Long, J As Long, Op(1) As Esprit.Operation, ReOrderOp(0) As Esprit.Operation
    Dim SortOrder(1) As EspritProperties.CustomProperty, SwapOps As Boolean
    With Document.Operations
        For i = 1 To .Count - 1
            For J = .Count To (i + 1) Step -1
                Set Op(0) = .Item(J - 1)
                Set SortOrder(0) = GetCustomLong(Op(0).CustomProperties, "SortOrder")
                If SortOrder(0) Is Nothing Then
                    SwapOps = True
                Else
                    Set Op(1) = .Item(J)
                    Set SortOrder(1) = GetCustomLong(Op(1).CustomProperties, "SortOrder")
                    If Not (SortOrder(1) Is Nothing) Then
                        If SortOrder(0).Value > SortOrder(1).Value Then SwapOps = True
                    End If
                End If
                If SwapOps Then
                    Set ReOrderOp(0) = Op(0)
                    Call .Reorder(ReOrderOp, J, False)
                    SwapOps = False
                End If
            Next
        Next
    End With
End Sub
Function getOperationByName(strOperationName As String) As Esprit.Operation
    Dim i As Long, Op As Esprit.Operation
    For Each Op In Document.Operations
        If (Op.Name = strOperationName) Then
            Set getOperationByName = Op
            Exit For
        End If
    Next
End Function
Sub ReorderOperation()
    Dim Op As Esprit.Operation
    Dim i As Long
    Dim strOperationName() As String
    
    Call Application.OutputWindow.Text("PartStockLength(Before): " & CStr(Document.LatheMachineSetup.PartStockLength) & vbCrLf)
    Document.LatheMachineSetup.PartStockLength = Round(GetCutOffXRightEnd(), 2)
    Call Application.OutputWindow.Text("PartStockLength(Updated): " & CStr(Document.LatheMachineSetup.PartStockLength) & vbCrLf)
    
    Call DeleteDummyOperation
    With Document.Operations
        ReDim strOperationName(1 To .Count)
        For i = 1 To .Count
            Set Op = .Item(i)
            Call SetCustomLong(Op.CustomProperties, "SortOrder", i)
            strOperationName(i) = .Item(i).Name
            'MsgBox ("A[" & CStr(I) & "]" & strOperationName(I))
        Next i
    End With
    
    Call SelectionSortStrings(strOperationName())
    For i = 1 To Document.Operations.Count
        If (UBound(filter(strOperationName, Op.Name)) > -1) Then
            'MsgBox ("B[" & CStr(I) & "]" & strOperationName(I))
            Call SetCustomLong(getOperationByName(strOperationName(i)).CustomProperties, "SortOrder", i)
        End If
    Next i

    Call RestoreLastSavedOperationOrder

End Sub
Public Sub SelectionSortStrings(ListArray() As String, _
                                Optional ByVal bAscending As Boolean = True, _
                                Optional ByVal bCaseSensitive As Boolean = False)
    
    Dim sSmallest       As String
    Dim lSmallest       As Long
    Dim lCount1         As Long
    Dim lCount2         As Long
    Dim lMin            As Long
    Dim lMax            As Long
    Dim lCompareType    As Long
    Dim lOrder          As Long
    
    lMin = LBound(ListArray)
    lMax = UBound(ListArray)
    
    If lMin = lMax Then
        Exit Sub
    End If
    
    ' Order Ascending or Descending?
    lOrder = IIf(bAscending, -1, 1)
    
    ' Case sensitive search or not?
    lCompareType = IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare)
    
    ' Loop through array swapping the smallest\largest (determined by lOrder)
    ' item with the current item
    For lCount1 = lMin To lMax - 1
        sSmallest = ListArray(lCount1)
        lSmallest = lCount1
        
        ' Find the smallest\largest item in the array
        For lCount2 = lCount1 + 1 To lMax
            If StrComp(ListArray(lCount2), sSmallest, lCompareType) = lOrder Then
                sSmallest = ListArray(lCount2)
                lSmallest = lCount2
            End If
        Next
        
        ' Just swap them, even if we are swapping it with itself,
        ' as it is generally quicker to do this than test first
        ' each time if we are already the smallest with a
        ' test like: If lSmallest <> lCount1 Then
        ListArray(lSmallest) = ListArray(lCount1)
        ListArray(lCount1) = sSmallest
    Next
End Sub

Public Sub showAdvancedNCCode()
    Call Document.GUI.NCCodeAdvanced(True)
End Sub

Public Sub CenterForm(ByVal frm As UserForm, Optional ByVal parent As UserForm = Nothing)
    '' Note: call this from frm's Load event!
    Dim r As Rectangle
    If parent Is Not Nothing Then
        r = parent.RectangleToScreen(parent.ClientRectangle)
    Else
        r = Screen.FromPoint(frm.Location).WorkingArea
    End If

    Dim x As Double
    x = r.Left + (r.Width - frm.Width) / 2
    Dim y As Double
    y = r.Top + (r.Height - frm.Height) / 2
    frm.Left = x
    frm.Top = y
End Sub

