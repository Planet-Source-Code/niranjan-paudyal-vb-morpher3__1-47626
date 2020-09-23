Attribute VB_Name = "ModFunctions"
'Part of VBMorpher3 written by Niranjan paudyal
'No parts of this program may be copied or used
'without contacting me first at
'nirpaudyal@hotmail.com (see the about box)
'If you like to use the avi file making module, or
'would like any help on that module, plese contact
'the author (contact address on mAVIDecs module)

Option Explicit
Option Base 1   'Set base to 1, insted of zero, it is a lot less confusing when handeling arrays
Type PointAPI
    x As Long
    y As Long
End Type

Type Triangle   'Holds the 3 vertices of triangles
    Vertex(1 To 3) As PointAPI
End Type

Type Grid
    filename As String          'full path of the file
    ControlPointRadius As Long  'the radius of the control point
    LineColor As Long           'color of the lines of the triangles
    ControlPointColor As Long   'color of control points
    GridDiamension As PointAPI  'the number of rectangle on the grid on X and Y axis (note that 2 triangles make 1 rectangle)
    GridPoint() As PointAPI     'the array holding the co-ordinates of each control point
    GridWidth As Long           'the width of grid in Pixels
    GridHeight As Long          'the height of grid in pixels
End Type

Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As PointAPI) As Long
Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As PointAPI, ByVal nCount As Long) As Long
Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Const HALFTONE = 4
Public Const COLORONCOLOR = 3

Public Function GetFileName(Path As String) As String
    Dim i As Long
    For i = Len(Path) To 1 Step -1
        If Mid(Path, i, 1) <> "\" Then
            GetFileName = Mid(Path, i, 1) & GetFileName
        Else
            Exit For
        End If
    Next i
End Function
Public Function GetPathName(Path As String) As String
    Dim i As Long
    For i = Len(Path) To 1 Step -1
        If Mid(Path, i, 1) = "\" Then
            Exit For
        End If
    Next i
    GetPathName = Mid(Path, 1, i)
    
    
End Function

Sub DrawGrid(P As PictureBox, G As Grid, Optional Zoom As Long = 1)
    Dim J As Long, i As Long, TmpPoint As PointAPI
    
    P.Cls
    
    'Draw the circle to indicate the control points of the grid
    'The control points  may be dragged about
    P.ForeColor = G.ControlPointColor
    For J = 1 To G.GridDiamension.y + 1
        For i = 1 To G.GridDiamension.x + 1
            Ellipse P.hdc, G.GridPoint(i, J).x * Zoom - G.ControlPointRadius, G.GridPoint(i, J).y * Zoom - G.ControlPointRadius, G.GridPoint(i, J).x * Zoom + G.ControlPointRadius, G.GridPoint(i, J).y * Zoom + G.ControlPointRadius
        Next i
    Next J
    
    'the lines
    P.ForeColor = G.LineColor
    For J = 1 To G.GridDiamension.y
        For i = 1 To G.GridDiamension.x
            MoveToEx P.hdc, G.GridPoint(i, J).x * Zoom, G.GridPoint(i, J).y * Zoom, TmpPoint
            LineTo P.hdc, G.GridPoint(i + 1, J).x * Zoom, G.GridPoint(i + 1, J).y * Zoom
            LineTo P.hdc, G.GridPoint(i + 1, J + 1).x * Zoom, G.GridPoint(i + 1, J + 1).y * Zoom
            LineTo P.hdc, G.GridPoint(i, J + 1).x * Zoom, G.GridPoint(i, J + 1).y * Zoom
            LineTo P.hdc, G.GridPoint(i, J).x * Zoom, G.GridPoint(i, J).y * Zoom
            LineTo P.hdc, G.GridPoint(i + 1, J + 1).x * Zoom, G.GridPoint(i + 1, J + 1).y * Zoom
        Next i
    Next J
    
    P.Refresh
End Sub

Sub LoadGrid(G As Grid)
    
    Dim Gw As Single, Gh As Single
    Dim i As Single, J As Single
    
    Gw = (G.GridWidth / G.GridDiamension.x) 'Get the horizontal size of the cells
    Gh = (G.GridHeight / G.GridDiamension.y) 'Get the vertical size of the cells
    
    'Redimension the grid size
    ReDim G.GridPoint(G.GridDiamension.x + 1, G.GridDiamension.y + 1)
    
    'Fill the grid array
    For J = 1 To G.GridDiamension.y
        For i = 1 To G.GridDiamension.x
            G.GridPoint(i, J).x = (i - 1) * Gw
            G.GridPoint(i, J).y = (J - 1) * Gh
        Next i
    Next J
    'fill the right side of the grid
    For J = 1 To G.GridDiamension.y
        G.GridPoint(G.GridDiamension.x + 1, J).x = G.GridWidth
        G.GridPoint(G.GridDiamension.x + 1, J).y = (J - 1) * Gh
    Next J
    'fill the bottom side of the grid
    For i = 1 To G.GridDiamension.x
        G.GridPoint(i, G.GridDiamension.y + 1).x = (i - 1) * Gw
        G.GridPoint(i, G.GridDiamension.y + 1).y = G.GridHeight
    Next i
    
    'the right-bottom point
        G.GridPoint(G.GridDiamension.x + 1, G.GridDiamension.y + 1).x = G.GridWidth
        G.GridPoint(G.GridDiamension.x + 1, G.GridDiamension.y + 1).y = G.GridHeight
    
    Dim T() As Triangle
    GenerateTriangleList G, T
End Sub

Sub GenerateTriangleList(G As Grid, ByRef Triangles() As Triangle)
    'this will create a list of triangles from the grids control point
    Dim i As Long, J As Long, Counter As Long
    ReDim Triangles(G.GridDiamension.x * G.GridDiamension.y * 2) 'this rediamensions the list with the number of triangles in the grid
    
    Counter = 1
    'Go through the grid getting the triangle vertices
    For J = 1 To G.GridDiamension.y
        For i = 1 To G.GridDiamension.x
            Triangles(Counter).Vertex(1) = G.GridPoint(i, J)
            Triangles(Counter).Vertex(2) = G.GridPoint(i, J + 1)
            Triangles(Counter).Vertex(3) = G.GridPoint(i + 1, J + 1)
            
            Triangles(Counter + 1).Vertex(1) = G.GridPoint(i, J)
            Triangles(Counter + 1).Vertex(2) = G.GridPoint(i + 1, J)
            Triangles(Counter + 1).Vertex(3) = G.GridPoint(i + 1, J + 1)
            Counter = Counter + 2
        Next i
    Next J
End Sub

Public Function WithinCircle(x As Long, y As Long, CircleX As Long, CircleY As Long, CircleRadius As Long) As Boolean
    'function to check if a point is within a circle
    If ((x - CircleX) * (x - CircleX) + (y - CircleY) * (y - CircleY)) < CircleRadius * CircleRadius Then WithinCircle = True
End Function
Sub FillPolygon(hdc As Long, a() As PointAPI, fillColor As Long, Count As Long)
    Dim hBrush As Long, Sel1 As Long
    Dim Poly As Long
    hBrush = CreateSolidBrush(fillColor)
    Sel1 = SelectObject(hdc, hBrush)
    Poly = Polygon(hdc, a(LBound(a)), Count)
    DeleteObject hBrush
    DeleteObject Sel1
    DeleteObject Poly
End Sub
Sub DrawTriangles(P As PictureBox, T() As Triangle)
    Dim i As Long
    For i = 1 To UBound(T)
        FillPolygon P.hdc, T(i).Vertex, i * 10, 3
    Next i
End Sub
Sub IncTriangle(T1() As Triangle, T2() As Triangle, DestT() As Triangle, FrameNo As Long, TotalFrames As Long)
    Dim i As Long, FrameRatio As Single
    FrameRatio = FrameNo / TotalFrames
    ReDim DestT(UBound(T1))
    For i = 1 To UBound(T1)
        With DestT(i)
            .Vertex(1).x = FrameRatio * (T2(i).Vertex(1).x - T1(i).Vertex(1).x) + T1(i).Vertex(1).x
            .Vertex(1).y = FrameRatio * (T2(i).Vertex(1).y - T1(i).Vertex(1).y) + T1(i).Vertex(1).y
            
            .Vertex(2).x = FrameRatio * (T2(i).Vertex(2).x - T1(i).Vertex(2).x) + T1(i).Vertex(2).x
            .Vertex(2).y = FrameRatio * (T2(i).Vertex(2).y - T1(i).Vertex(2).y) + T1(i).Vertex(2).y
            
            .Vertex(3).x = FrameRatio * (T2(i).Vertex(3).x - T1(i).Vertex(3).x) + T1(i).Vertex(3).x
            .Vertex(3).y = FrameRatio * (T2(i).Vertex(3).y - T1(i).Vertex(3).y) + T1(i).Vertex(3).y
        End With
    Next i
End Sub
Sub WrapPictures(Source1 As BITMAPINFO, Source1T() As Triangle, Source2 As BITMAPINFO, Source2T() As Triangle, MidTriangles As BITMAPINFO, MidT() As Triangle, FrameRatio As Single)
    Dim x As Long, y As Long, Index As Single, LastIndex As Long
    Dim U As Single, V As Single, W As Single
    Dim PicW As Long, PicH As Long
    Dim TriangleArea As Single, Final1 As PointAPI, Final2 As PointAPI
    Dim FrameRatio1 As Single
    
    PicW = MidTriangles.bmiHeader.biWidth - 1
    PicH = -MidTriangles.bmiHeader.biHeight - 1
    
    FrameRatio1 = 1 - FrameRatio
    
    For y = 0 To PicH
        For x = 0 To PicW
            
            Index = RGB(MidTriangles.pBits(2, x, y), MidTriangles.pBits(1, x, y), MidTriangles.pBits(0, x, y)) * 0.1
        
        
            If Index <= UBound(MidT) Then
                If Index <> LastIndex Then TriangleArea = GetTriangleArea(MidT(Index).Vertex(1), MidT(Index).Vertex(2), MidT(Index).Vertex(3))
                
                PointToB MidT(Index), x, y, TriangleArea, U, V, W
                
               'convert B to point
                With Source1T(Index)
                    Final1.x = U * .Vertex(1).x + V * .Vertex(2).x + W * .Vertex(3).x
                    Final1.y = U * .Vertex(1).y + V * .Vertex(2).y + W * .Vertex(3).y
                End With
                
                PointToB MidT(Index), x, y, TriangleArea, U, V, W
                With Source2T(Index)
                    Final2.x = U * .Vertex(1).x + V * .Vertex(2).x + W * .Vertex(3).x
                    Final2.y = U * .Vertex(1).y + V * .Vertex(2).y + W * .Vertex(3).y
                End With

                
                If Final1.x > PicW Then Final1.x = PicW
                If Final1.y > PicH Then Final1.y = PicH
                If Final1.x < 0 Then Final1.x = 0
                If Final1.y < 0 Then Final1.y = 0
                
                If Final2.x > PicW Then Final2.x = PicW
                If Final2.y > PicH Then Final2.y = PicH
                If Final2.x < 0 Then Final2.x = 0
                If Final2.y < 0 Then Final2.y = 0
                
                MidTriangles.pBits(0, x, y) = (FrameRatio1 * Source1.pBits(0, Final1.x, Final1.y)) + (FrameRatio * Source2.pBits(0, Final2.x, Final2.y))
                MidTriangles.pBits(1, x, y) = (FrameRatio1 * Source1.pBits(1, Final1.x, Final1.y)) + (FrameRatio * Source2.pBits(1, Final2.x, Final2.y))
                MidTriangles.pBits(2, x, y) = (FrameRatio1 * Source1.pBits(2, Final1.x, Final1.y)) + (FrameRatio * Source2.pBits(2, Final2.x, Final2.y))
                LastIndex = Index
            
            End If
        Next x
    Next y
End Sub
'converts a point to Barycentric coordinates
Sub PointToB(T As Triangle, x, y, TriArea As Single, ByRef U As Single, ByRef V As Single, ByRef W As Single)
    Dim PPt As PointAPI, tmpArea As Single
    Dim M As Single, Bm As Single
    Dim vX As Long, vY As Long
    
    PPt.x = x: PPt.y = y
    If TriArea = 0 Then tmpArea = 5E+20 Else tmpArea = 0.5 / TriArea
    
    vX = PPt.x - T.Vertex(2).x
    vY = PPt.y - T.Vertex(2).y
    If vX <> 0 Then M = vY / vX Else M = 10000000
    Bm = Abs(-M * T.Vertex(3).x + T.Vertex(3).y - PPt.y + M * PPt.x) / Sqr(M * M + 1)
    U = (Sqr(vX * vX + vY * vY) * Bm) * tmpArea
        
    vX = T.Vertex(1).x - PPt.x
    vY = T.Vertex(1).y - PPt.y
    If vX <> 0 Then M = vY / vX Else M = 1E+20
    Bm = Abs(-M * T.Vertex(3).x + T.Vertex(3).y - T.Vertex(1).y + M * T.Vertex(1).x) / Sqr(M * M + 1)
    V = (Sqr(vX * vX + vY * vY) * Bm) * tmpArea
    
    W = 1 - U - V

End Sub
'converts Barycentric coordinates to a point
Function BtoPoint(U As Single, V As Single, W As Single, T As Triangle) As PointAPI
    With T
        BtoPoint.x = U * .Vertex(1).x + V * .Vertex(2).x + W * .Vertex(3).x
        BtoPoint.y = U * .Vertex(1).y + V * .Vertex(2).y + W * .Vertex(3).y
    End With
End Function
'function to calculate the area of a triangle
Function GetTriangleArea(a As PointAPI, B As PointAPI, c As PointAPI) As Single
    Dim M As Single, Bm As Single
    Dim vX As Long, vY As Long
    
    vX = a.x - B.x
    vY = a.y - B.y
    If a.x <> B.x Then
        M = vY / vX
    Else
        M = 1E+20
    End If
    Bm = Abs(-M * c.x + c.y - a.y + M * a.x) / Sqr(M * M + 1)
    GetTriangleArea = 0.5 * Sqr(vX * vX + vY * vY) * Bm
End Function

Function SaveMorphFile(OutputFile As String, ByRef G1 As Grid, ByRef G2 As Grid, ByVal OutPutPath As String, ByVal TotalFrames As Long, ByRef SaveAsBMP As Boolean, ByRef FPS As Long) As String
    Dim i As Long, J As Long
    Dim Fn
    On Error GoTo Ex1
    Open OutputFile For Output As #1
    'write output directory and total frames
    Write #1, OutPutPath
    Write #1, TotalFrames
    'Write the grid 1 data to file
    Write #1, G1.ControlPointColor
    Write #1, G1.ControlPointRadius
    Write #1, G1.filename
    Write #1, G1.GridDiamension.x
    Write #1, G1.GridDiamension.y
    Write #1, G1.GridHeight
    For J = 1 To G1.GridDiamension.y + 1
        For i = 1 To G1.GridDiamension.x + 1
            Write #1, G1.GridPoint(i, J).x ', G1.GridPoint(I, J).Y
            Write #1, G1.GridPoint(i, J).y
        Next i
    Next J
    Write #1, G1.GridWidth
    Write #1, G1.LineColor
    'Write the grid 2 data to file
    Write #1, G2.ControlPointColor
    Write #1, G2.ControlPointRadius
    Write #1, G2.filename
    Write #1, G2.GridDiamension.x
    Write #1, G2.GridDiamension.y
    Write #1, G2.GridHeight
    For J = 1 To G2.GridDiamension.y + 1
        For i = 1 To G2.GridDiamension.x + 1
            Write #1, G2.GridPoint(i, J).x
            Write #1, G2.GridPoint(i, J).y
        Next i
    Next J
    Write #1, G2.GridWidth
    Write #1, G2.LineColor
    Write #1, SaveAsBMP
    Write #1, FPS
    Close #1
    Exit Function
Ex1:
    SaveMorphFile = "There was a problem while creating '" & OutputFile & "', file will not be saved."
    Close #1
    Exit Function
End Function

Function LoadMorphFile(filename As String, ByRef G1 As Grid, ByRef G2 As Grid, ByRef OutPutPath As String, ByRef TotalFrames As Long, ByRef SaveAsBMP As Boolean, ByRef FPS As Long) As String
    Dim i As Long, J As Long, Var As String
    Open filename For Input As #1
    
    Line Input #1, Var
    Var = IIf(left(Var, 1) = """", Mid(Var, 2, Len(Var) - 2), Var)
    If Dir(Var, vbDirectory) <> "" Then OutPutPath = Var Else OutPutPath = App.Path & "\Morphed\"
    Line Input #1, Var: TotalFrames = Var
    'Get grid 1 data
    Line Input #1, Var: G1.ControlPointColor = Var
    Line Input #1, Var: G1.ControlPointRadius = Var
    Line Input #1, Var
    Var = IIf(left(Var, 1) = """", Mid(Var, 2, Len(Var) - 2), Var)
    
    If Dir(Var) = "" Then
        If Dir(GetPathName(filename) & GetFileName(Var)) <> "" Then
            G1.filename = GetPathName(filename) & GetFileName(Var)
        Else
            LoadMorphFile = "There was an problem when opening the picture '" & Var & "'. Make sure it exists."
            Close #1
            Exit Function
        End If
    Else
        G1.filename = Var
    End If
    Line Input #1, Var: G1.GridDiamension.x = Var
    Line Input #1, Var: G1.GridDiamension.y = Var
    Line Input #1, Var: G1.GridHeight = Var
    ReDim G1.GridPoint(G1.GridDiamension.x + 1, G1.GridDiamension.y + 1)
    
    For J = 1 To G1.GridDiamension.y + 1
        For i = 1 To G1.GridDiamension.x + 1
            Line Input #1, Var: G1.GridPoint(i, J).x = Var
            Line Input #1, Var: G1.GridPoint(i, J).y = Var
        Next i
    Next J
    Line Input #1, Var:  G1.GridWidth = Var
    Line Input #1, Var:  G1.LineColor = Var
    'Get grid 2 data
    Line Input #1, Var:  G2.ControlPointColor = Var
    Line Input #1, Var:  G2.ControlPointRadius = Var
    Line Input #1, Var
    Var = IIf(left(Var, 1) = """", Mid(Var, 2, Len(Var) - 2), Var)

    If Dir(Var) = "" Then

        If Dir(GetPathName(filename) & GetFileName(Var)) <> "" Then
            G2.filename = GetPathName(filename) & GetFileName(Var)
        Else
            LoadMorphFile = "There was an problem when opening the picture '" & Var & "'. Make sure it exists."
            Close #1
            Exit Function
        End If
    Else
        G2.filename = Var
    End If
    Line Input #1, Var:  G2.GridDiamension.x = Var
    Line Input #1, Var:  G2.GridDiamension.y = Var
    Line Input #1, Var:  G2.GridHeight = Var
    ReDim G2.GridPoint(G2.GridDiamension.x + 1, G2.GridDiamension.y + 1)
    
    For J = 1 To G2.GridDiamension.y + 1
        For i = 1 To G2.GridDiamension.x + 1
            Line Input #1, Var:  G2.GridPoint(i, J).x = Var
            Line Input #1, Var:  G2.GridPoint(i, J).y = Var
        Next i
    Next J
    Line Input #1, Var: G2.GridWidth = Var
    Line Input #1, Var: G2.LineColor = Var
    Line Input #1, Var: SaveAsBMP = Var
    Line Input #1, Var: FPS = Var
    Close #1
    Exit Function
Ex:
    LoadMorphFile = "There was an problem when opening '" & filename & "'."
    Close #1
End Function
