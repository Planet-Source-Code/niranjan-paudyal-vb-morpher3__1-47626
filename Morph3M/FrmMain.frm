VERSION 5.00
Begin VB.Form FrmMain 
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   3870
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4170
   ScaleWidth      =   3870
   Begin VB.PictureBox Poriginal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   2
      Left            =   1680
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   10
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Poriginal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   1
      Left            =   120
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Pb 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.PictureBox cBlock 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   125
         TabIndex        =   6
         Top             =   2280
         Width           =   1875
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "FrmMain.frx":000C
            Left            =   1035
            List            =   "FrmMain.frx":0019
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   0
            Width           =   870
         End
         Begin VB.PictureBox picCol 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   30
            Picture         =   "FrmMain.frx":002F
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   65
            TabIndex        =   7
            Top             =   30
            Width           =   975
         End
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   2280
         Width           =   1575
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1695
         Left            =   3360
         TabIndex        =   1
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox P 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   720
         Index           =   2
         Left            =   1200
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   44
         TabIndex        =   4
         Top             =   120
         Width           =   660
      End
      Begin VB.PictureBox P 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   720
         Index           =   1
         Left            =   240
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   52
         TabIndex        =   3
         Top             =   120
         Width           =   780
      End
      Begin VB.PictureBox MorphT 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         DrawStyle       =   5  'Transparent
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   2160
         ScaleHeight     =   53
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   49
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewMorph 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpenMorph 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSaveMorph 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuCloseMorph 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Option"
      Begin VB.Menu mnuMorphEditorOption 
         Caption         =   "Morph options..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMorphIt 
         Caption         =   "Morph it!"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Part of VBMorpher3 written by Niranjan paudyal
'No parts of this program may be copied or used
'without contacting me first at
'nirpaudyal@hotmail.com (see the about box)
'If you like to use the avi file making module, or
'would like any help on that module, plese contact
'the author (contact address on mAVIDecs module)

Option Explicit
Option Base 1   'Set base to 1, insted of zero, _
                it is a lot less confusing when handeling arrays

Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As PointAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWND As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim G(2) As Grid   'This holds the data requried for the triangle, for the two grids of two pictures
Dim OutPutPath As String    'Location of the file output
Dim SaveBMP As Boolean  ' is the files to be saved as BMP or AVI
Dim CurrentPointDragged As PointAPI    'the index of the pointDragged otherwise, -1
Dim FPS As Long     'FPS of avi file
Dim LastPoint As PointAPI   'The location of a control point before it was dragged
Dim TotalFrames As Long 'The total number of frames for the morph
Dim openedFileName As String  'Name of the current file, otherwise ""
Dim MorphDirty As Boolean   'Has the morph been changed?
Dim ZoomValue As Long   'The zoom ammount


Public Sub OpenMorph(filename As String)
    Me.Visible = False
    Dim RetVal As String
    
    'This sub is only required to open a file from MDIFormMain
    
    RetVal = ModFunctions.LoadMorphFile(filename, G(1), G(2), OutPutPath, TotalFrames, SaveBMP, FPS)
    
    If RetVal <> "" Then
        MsgBox RetVal, vbCritical, "Error"
        Unload Me
        Exit Sub
    End If
    
    On Error GoTo Er1
    P(1).Picture = LoadPicture(G(1).filename)
    Poriginal(1).Picture = LoadPicture(G(1).filename)
    On Error GoTo Er2
    P(2).Picture = LoadPicture(G(2).filename)
    Poriginal(2).Picture = LoadPicture(G(2).filename)
    
    DrawGrid P(1), G(1), ZoomValue
    DrawGrid P(2), G(2), ZoomValue
    openedFileName = filename
    Me.Caption = filename
    Me.Visible = True
    Me.Show
    
    AlignScroll
    Exit Sub
Er1:
    MsgBox "The morph file makes reference to '" & G(1).filename & "', it seems to contains an error. Morph will not be loaded", vbExclamation, "Error!"
    Unload Me
    Exit Sub
Er2:
    MsgBox "The morph file makes reference to '" & G(2).filename & "', it seems to contains an error. Morph will not be loaded", vbExclamation, "Error!"
    Unload Me
End Sub
Public Sub NewMorph(FileName1 As String, FileName2 As String)
    
    On Error GoTo Er1
    P(1).Picture = LoadPicture(FileName1)
    Poriginal(1).Picture = LoadPicture(FileName1)
    On Error GoTo Er2
    P(2).Picture = LoadPicture(FileName2)
    Poriginal(2).Picture = LoadPicture(FileName2)
    
    TotalFrames = 19
    OutPutPath = App.Path & "\Morphed\"
    If Dir(OutPutPath, vbDirectory) = "" Then MkDir (OutPutPath)
    CurrentPointDragged.x = -1
    CurrentPointDragged.y = -1
    G(1).LineColor = vbWhite
    G(1).ControlPointColor = vbRed
    G(1).GridDiamension.x = 5
    G(1).GridDiamension.y = 5
    G(1).ControlPointRadius = 5
    G(1).GridWidth = P(1).ScaleWidth
    G(1).GridHeight = P(1).ScaleHeight
    G(2) = G(1)
    G(1).filename = FileName1
    G(2).filename = FileName2
    LoadGrid G(1)
    DrawGrid P(1), G(1), ZoomValue
    LoadGrid G(2)
    DrawGrid P(2), G(2), ZoomValue
    Me.Show
    AlignScroll
    SaveBMP = True
    FPS = 10
    Exit Sub

Er1:
    MsgBox "There was a problem loading '" & FileName1 & "'. It might not exits or it contains an error. Morph will not be loaded", vbExclamation, "Error!"
    Unload Me
    Exit Sub
Er2:
    MsgBox "There was a problem loading '" & FileName2 & "'. It might not exits or it contains an error. Morph will not be loaded", vbExclamation, "Error!"
    Unload Me
End Sub

Private Sub AlignScroll()
    On Error Resume Next
    Dim Hbar As Boolean, Vbar As Boolean
    Dim Tw As Long
    If HScroll1.Value > 0 Then HScroll1.Value = 0
    If VScroll1.Value > 0 Then VScroll1.Value = 0
    
    P(1).Move 0, 0
    P(2).Move P(1).ScaleWidth + 10, 0
    Pb.Move 0, 0, ScaleWidth, ScaleHeight
    
    If P(2).left + P(2).ScaleWidth > Pb.ScaleWidth - VScroll1.Width Then Hbar = True
    If P(1).Height > Pb.ScaleHeight - HScroll1.Height Then Vbar = True
    
    cBlock.Move Pb.ScaleWidth - cBlock.ScaleWidth, Pb.ScaleHeight - cBlock.Height ' VScroll1.Height ', HScroll1.Width, VScroll1.Height, cBlock.Width
    VScroll1.Move Pb.ScaleWidth - VScroll1.Width, 0, VScroll1.Width, cBlock.top
    HScroll1.Move 0, cBlock.top, cBlock.left, cBlock.ScaleHeight
    
    If Hbar Then
        HScroll1.Max = (P(2).left + P(2).ScaleWidth) - (Pb.ScaleWidth - VScroll1.Width)
        HScroll1.LargeChange = HScroll1.Max / 5
    End If
    
    If Vbar Then
        VScroll1.Max = P(1).Height - (Pb.ScaleHeight - HScroll1.Height)
        VScroll1.LargeChange = VScroll1.Max / 5
    End If
    HScroll1.Enabled = Hbar
    VScroll1.Enabled = Vbar
    
    HScroll1.Visible = (cBlock.left >= 0)

End Sub

Private Sub Combo1_Click()
    Dim TempDir As String
    TempDir = IIf(right(App.Path, 1) = "\", App.Path, App.Path & "\")

    If ZoomValue = 0 Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    Dim Dib1 As BITMAPINFO, Dib2 As BITMAPINFO
    
    MakePicMem Poriginal(1), Dib1, 1
    MakePicMem Poriginal(2), Dib2, 1
    
    ZoomValue = Mid(Combo1.Text, 1, Len(Combo1.Text) - 1) / 100
    P(1).Picture = LoadPicture("")
    P(2).Picture = LoadPicture("")
    
    P(1).Width = Poriginal(1).ScaleWidth * ZoomValue
    P(1).Height = Poriginal(1).ScaleHeight * ZoomValue
    P(2).Width = Poriginal(2).ScaleWidth * ZoomValue
    P(2).Height = Poriginal(2).ScaleHeight * ZoomValue
    
    StretchDIBits P(1).hdc, 0, 0, P(1).ScaleWidth, P(1).ScaleHeight, 0, 0, Dib1.bmiHeader.biWidth - 1, -Dib1.bmiHeader.biHeight - 1, Dib1.pBits(0, 0, 0), Dib1, 1, vbSrcCopy
    SavePicture P(1).Image, TempDir & "tmp.bmp"
    P(1).Picture = LoadPicture(TempDir & "tmp.bmp")
    
    StretchDIBits P(2).hdc, 0, 0, P(2).ScaleWidth, P(2).ScaleHeight, 0, 0, Dib2.bmiHeader.biWidth - 1, -Dib2.bmiHeader.biHeight - 1, Dib2.pBits(0, 0, 0), Dib2, 1, vbSrcCopy
    SavePicture P(2).Image, TempDir & "tmp.bmp"
    P(2).Picture = LoadPicture(TempDir & "tmp.bmp")
    
    P(2).Move P(1).ScaleWidth + 10
    
    DrawGrid P(1), G(1), ZoomValue
    DrawGrid P(2), G(2), ZoomValue
    
    AlignScroll
    
    MakePicMem Poriginal(1), Dib1, 3
    MakePicMem Poriginal(2), Dib2, 3
    Me.MousePointer = vbArrow
End Sub

Private Sub Command1_Click()
P(1).Cls
End Sub

Private Sub mnuAbout_Click()
    FrmAbout.Show 1
End Sub
Private Sub picCol_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    picCol_MouseMove Button, Shift, x, y
End Sub

Private Sub picCol_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    If x > picCol.ScaleWidth - 1 Then x = picCol.ScaleWidth - 1
    If y > picCol.ScaleHeight - 1 Then y = picCol.ScaleHeight - 1
    If Button = 1 Then
        G(1).LineColor = picCol.Point(x, y)
        G(2).LineColor = picCol.Point(x, y)
        DrawGrid P(1), G(1), ZoomValue
        DrawGrid P(2), G(2), ZoomValue
        'indicate that morph has been changed
        MorphDirty = True
    End If
    If Button = 2 Then
        G(1).ControlPointColor = picCol.Point(x, y)
        G(2).ControlPointColor = picCol.Point(x, y)
        DrawGrid P(1), G(1), ZoomValue
        DrawGrid P(2), G(2), ZoomValue
        'indicate that morph has been changed
        MorphDirty = True
    End If
End Sub
Private Sub Form_Activate()
    CurrentPointDragged.x = -1
    CurrentPointDragged.y = -1
End Sub

Private Sub Form_Load()
    Dim i As Long
    ZoomValue = 0
    Combo1.Text = "100%"
    ZoomValue = 1
    Me.Show
    AlignScroll
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim RetVal As Long
    If MorphDirty Then
        RetVal = MsgBox("'" & Me.Caption & "' has been changed, save changes?", vbYesNoCancel, "Save?")
        If RetVal = vbCancel Then Cancel = 1
        If RetVal = vbYes Then
            mnuSaveMorph_Click
        End If
    End If
End Sub

Private Sub Form_Resize()
    AlignScroll
End Sub

Private Sub HScroll1_Change()
    P(1).left = -HScroll1.Value
    P(2).left = P(1).left + P(1).ScaleWidth + 10
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

Private Sub mnuCloseMorph_Click()
    Unload Me
End Sub

Public Sub mnuMorphEditorOption_Click()
    If Dir(OutPutPath, vbDirectory) = "" Then MkDir (OutPutPath)
    FrmOptions.Dir1 = OutPutPath
    FrmOptions.Text1 = TotalFrames + 1
    FrmOptions.Text2 = G(1).GridDiamension.x
    FrmOptions.Text3 = G(1).GridDiamension.y
    FrmOptions.Text4 = G(1).ControlPointRadius
    FrmOptions.Option1(IIf(SaveBMP = True, 0, 1)).Value = True
    FrmOptions.Text5.Enabled = Not SaveBMP
    FrmOptions.Label5.Enabled = Not SaveBMP
    FrmOptions.Text5 = FPS
    FrmOptions.Show 1
    'If the user didnt press the  cancel butten then...
    If FrmOptions.Cancel = False Then
        'check to see if the grid size has been changed...
        If G(1).GridDiamension.x <> FrmOptions.GridX Or G(1).GridDiamension.y <> FrmOptions.GridY Then
            'if there are far too many triangles for the pictures then
            If (Poriginal(1).ScaleWidth / FrmOptions.GridX) < 5 Or (Poriginal(1).ScaleHeight / FrmOptions.GridY) < 5 Then
                MsgBox "You have set the value of the grid width or the grid height too much, the triangles will not fit onto the image!", vbExclamation, "Too many triangles."
                GoTo Ex
            End If
                
            'if it was changed then conform with user
            If MsgBox("You chose to change the grid diamensions, this will reset all control points, continue?", vbQuestion Or vbYesNo, "Grid changed.") = vbYes Then
            'if yes selected
                'resize the grids
                G(1).GridDiamension.x = FrmOptions.GridX
                G(1).GridDiamension.y = FrmOptions.GridY
                G(2).GridDiamension.x = FrmOptions.GridX
                G(2).GridDiamension.y = FrmOptions.GridY
                'reset the grids
                LoadGrid G(1)
                LoadGrid G(2)
                'redraw the grids
                DrawGrid P(1), G(1), ZoomValue
                DrawGrid P(2), G(2), ZoomValue
            End If
        End If
Ex:
        OutPutPath = IIf(right(FrmOptions.MorphPath, 1) = "\", FrmOptions.MorphPath, FrmOptions.MorphPath & "\")
        TotalFrames = FrmOptions.NoFrames
        SaveBMP = FrmOptions.SaveBMPformat
        FPS = FrmOptions.FPS
        If G(1).ControlPointRadius <> FrmOptions.CPRadius Then
            G(1).ControlPointRadius = FrmOptions.CPRadius
            G(2).ControlPointRadius = FrmOptions.CPRadius
            DrawGrid P(1), G(1), ZoomValue
            DrawGrid P(2), G(2), ZoomValue
        End If
        'indicate that morph has been changed
        MorphDirty = True
    End If
End Sub

Private Sub mnuMorphIt_Click()
    Dim InitalTitle As String   'the title of the form
    'This is the main sub where all the functions to morph _
    the images and to save it is called.
    
    'First, if the output directory exists, delete it and all its bmp contents, then make _
    otherwise, just create the directory
    If Dir(OutPutPath, vbDirectory) <> "" Then
    On Error GoTo DirDeleteError:
        If Dir(OutPutPath & "*.bmp", vbNormal) <> "" Then Kill (OutPutPath & "*.bmp")
    Else
    On Error GoTo DirCreateError:
        MkDir (OutPutPath)
    End If
    
    
    On Error GoTo UnknownError:
    Me.MousePointer = vbHourglass
    'now hide the two images to be morphed and shop the pricture _
    box in wich the images are morphed
    P(1).Visible = False
    P(2).Visible = False

    MorphT.Picture = LoadPicture("")    'clear the morphing picture box of any images
    MorphT.Move 0, 0, Poriginal(1).ScaleWidth, Poriginal(1).ScaleHeight  'Reposition+ resize it
    MorphT.Visible = True   'show it
    Pb.Refresh
    
    'Now everything is ready, we need to morph the pictures now!
    Dim MorphedPicture As BITMAPINFO    'This holds the bits for the morphed frame
    Dim S1 As BITMAPINFO                'This holds the bits for the frist picture
    Dim S2 As BITMAPINFO                'This holds the bits for the second picture
    Dim CurT() As Triangle              'This holds the triangles in the grid for the current frame
    Dim T1() As Triangle                'This holds the triangles in the first grid
    Dim T2() As Triangle                'This holds the triangls in the second grid
    Dim i As Long                       'Frame counter for going through all frames
    Dim RetVal As Long
    
    InitalTitle = Me.Caption
    
    'First create a list of triangles from each ot the two grids
    GenerateTriangleList G(1), T1
    GenerateTriangleList G(2), T2
    
    'Clear the source picture boxes, note that this deletes _
    only the lines, not the source pictures
    P(1).Cls
    P(2).Cls
    'Now create DIB from the two pictures
    ModDib.MakePicMem Poriginal(1), S1, 1
    ModDib.MakePicMem Poriginal(2), S2, 1
    
    'Go through each frame
    For i = 0 To TotalFrames
        Me.Caption = i + 1 & "/" & TotalFrames + 1
        'This will incriment the triangles from the first grid to _
        the one on the second grid completly
        IncTriangle T1, T2, CurT, i, TotalFrames
        'Draw the triangles onto the Morph picture box
        DrawTriangles MorphT, CurT
        'Create a DIB of this picture
        ModDib.MakePicMem MorphT, MorphedPicture, 1
        'Wrap the picture
        WrapPictures S1, T1, S2, T2, MorphedPicture, CurT, CSng(i / TotalFrames)
        
        'Set the DIB onto the morph picture box
        ModDib.MakePicMem MorphT, MorphedPicture, 2
        
        'save the picture
        SavePicture Me.MorphT.Image, IIf(right(OutPutPath, 1) = "\", OutPutPath, OutPutPath & "\") & i & ".bmp"
        
        'Delete the DIB
        ModDib.MakePicMem MorphT, MorphedPicture, 3
        'Refresh the picture box
        MorphT.Refresh
    Next i
    
    'Delete the Source pictures DIBs
    ModDib.MakePicMem Poriginal(1), S1, 3
    ModDib.MakePicMem Poriginal(2), S2, 3
    
    'Redraw the grid
    DrawGrid P(2), G(2), ZoomValue
    DrawGrid P(1), G(1), ZoomValue
    'Show the picture boxes
    P(1).Visible = True
    P(2).Visible = True
    
    'Hide the morph picture box
    MorphT.Visible = False

    'if we are to save the files in AVI format then
    If SaveBMP = False Then
        Dim RetString As String, AVIFileName As String, N As Long
        'get a name for the avi file
        If openedFileName <> "" Then
            For N = Len(openedFileName) - 4 To 1 Step -1
                If Mid(openedFileName, N, 1) <> "\" Then
                    AVIFileName = Mid(openedFileName, N, 1) & AVIFileName
                Else
                    Exit For
                End If
            Next N
            AVIFileName = AVIFileName & ".avi"
        Else
            AVIFileName = "Output.avi"
        End If
        Me.Caption = "Writing AVI file"
        'This next function will return a error discription if any, or will return nothing otherwise.
        RetString = nWriteAvi(OutPutPath, OutPutPath & AVIFileName, FPS, Me.HWND)
        'if there was an error display it
        If RetString <> "" And RetString <> "user_cancel" Then MsgBox RetString, vbCritical, "Error"
        'delete all the bmp files
        Kill OutPutPath & "*.bmp"
    End If
            
    AlignScroll
    Me.Caption = InitalTitle
    Me.MousePointer = vbArrow
    
    'if morphing was a complete success then...
    If RetString = "" Then
        RetVal = MsgBox("Morphing was completed successfully, would you like to open the folder containing the morphed file(s)?", vbQuestion Or vbYesNo, "Success!")
        If RetVal = vbYes Then ShellExecute Me.HWND, "open", OutPutPath, vbNull, vbNull, 1
    End If
    
    Exit Sub
DirCreateError:
    MsgBox "The directory '" & OutPutPath & "' could not be created, morphing will not progress.", vbExclamation, "Error"
    AlignScroll
    Exit Sub
DirDeleteError:
    MsgBox "The contents of the directory '" & OutPutPath & "' could not be deleted. Check that the files are not open. Morphing will not progress.", vbExclamation, "Error"
    AlignScroll
    Exit Sub
UnknownError:
    MsgBox "There was an unknown error creating a morph files!", vbExclamation, "Error"
    'Delete the Source pictures DIBs
    ModDib.MakePicMem Poriginal(1), S1, 3
    ModDib.MakePicMem Poriginal(2), S2, 3
    
    'Redraw the grid
    DrawGrid P(2), G(2), ZoomValue
    DrawGrid P(1), G(1), ZoomValue
    'Show the picture boxes
    P(1).Visible = True
    P(2).Visible = True
    
    'Hide the morph picture box
    MorphT.Visible = False
    
    AlignScroll
    Me.Caption = InitalTitle
    Me.MousePointer = vbArrow
End Sub

Private Sub mnuNewMorph_Click()
    MDIFormMain.mnuNewMorph_Click
End Sub

Private Sub mnuOpenMorph_Click()
    MDIFormMain.mnuOpenMorph_Click
End Sub

Private Sub mnuSaveMorph_Click()
    Dim CD As cFileDlg, FileSave As String
    
    'if the file was already opened, and the file exists then...
    If openedFileName <> "" And Dir(openedFileName, vbNormal) <> "" Then
        SaveMorphFile openedFileName, G(1), G(2), OutPutPath, TotalFrames, SaveBMP, FPS
        MorphDirty = False
    Else
        Set CD = New cFileDlg
        With CD
            .DlgTitle = "Save morph."
            .Filter = "Morph2 file (*.mrf)|*.mrf"
            .OverwritePrompt = True
            .OwnerHwnd = Me.HWND
            .InitDirectory = CurDir
        End With
        
        If CD.VBGetSaveFileName(FileSave) <> False Then
            SaveMorphFile IIf(right(FileSave, 4) = ".mrf", FileSave, FileSave & ".mrf"), G(1), G(2), OutPutPath, TotalFrames, SaveBMP, FPS
            openedFileName = IIf(right(FileSave, 4) = ".mrf", FileSave, FileSave & ".mrf")
            Me.Caption = openedFileName
            MorphDirty = False
        End If
        
        Set CD = Nothing
    End If
    
    
End Sub

Private Sub P_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long, J As Long, SelIndex As Long, hBrush As Long
    If Button = 1 Then
        'go through all points
        For J = 1 To G(Index).GridDiamension.y + 1
            For i = 1 To G(Index).GridDiamension.x + 1
                'If the location where the mouse was pressed is within _
                any of the control points then...
                If WithinCircle(CLng(x), CLng(y), G(Index).GridPoint(i, J).x * ZoomValue, G(Index).GridPoint(i, J).y * ZoomValue, G(Index).ControlPointRadius) Then
                    CurrentPointDragged.x = i
                    CurrentPointDragged.y = J
                    'obtain the location of this point, this may be used _
                    on the MouseUp event to restore the point to its original _
                    location if the point is dropped in an invalid location
                    
                    LastPoint = G(Index).GridPoint(i, J)
                    
                    'get the index of the picture box opposit to the selected box
                    SelIndex = IIf(Index = 1, 2, 1)
                    hBrush = CreateSolidBrush(G(SelIndex).LineColor) 'create a brush
                    SelectObject P(SelIndex).hdc, hBrush 'set the brush
                    'draw a circle on the opposit picture box _
                    this will indicate which corrsponding contorl point _
                    is currently selected
                    Ellipse P(SelIndex).hdc, G(SelIndex).GridPoint(i, J).x * ZoomValue - G(SelIndex).ControlPointRadius, G(SelIndex).GridPoint(i, J).y * ZoomValue - G(SelIndex).ControlPointRadius, G(SelIndex).GridPoint(i, J).x * ZoomValue + G(SelIndex).ControlPointRadius, G(SelIndex).GridPoint(i, J).y * ZoomValue + G(SelIndex).ControlPointRadius
                    'clear mem
                    DeleteObject hBrush
                    P(SelIndex).Refresh
                    'no need to continue as the control point has been selected _
                    so just exit
                    Exit Sub
                End If
            Next i
        Next J
        'If the loop has been completed, a point was not selected _
        so indicate this by setting values to -1
        CurrentPointDragged.x = -1
        CurrentPointDragged.y = -1
    End If

End Sub

Private Sub P_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If CurrentPointDragged.x <> -1 Then
            'make sure that the points tragged is within the picture box
            If x < 0 Then x = 0
            If y < 0 Then y = 0
            If x > P(Index).ScaleWidth Then x = P(Index).ScaleWidth
            If y > P(Index).ScaleHeight Then y = P(Index).ScaleHeight
            
            'If they are the edge points, make sure that they are dragged _
            only along the edge of the picture box
            If CurrentPointDragged.x = 1 Then x = 0
            If CurrentPointDragged.y = 1 Then y = 0
            If CurrentPointDragged.x = G(Index).GridDiamension.x + 1 Then x = P(Index).ScaleWidth 'G(Index).GridPoint(CurrentPointDragged.X, CurrentPointDragged.Y).X
            If CurrentPointDragged.y = G(Index).GridDiamension.y + 1 Then y = P(Index).ScaleHeight 'G(Index).GridPoint(CurrentPointDragged.X, CurrentPointDragged.Y).Y
            
            'Update grid point array and draw the grid
            G(Index).GridPoint(CurrentPointDragged.x, CurrentPointDragged.y).x = x / ZoomValue
            G(Index).GridPoint(CurrentPointDragged.x, CurrentPointDragged.y).y = y / ZoomValue
            DrawGrid P(Index), G(Index), ZoomValue
        End If
    End If
End Sub

Private Sub P_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Dim SelIndex As Long
        'if a point was selected then
        If CurrentPointDragged.x <> -1 Then
            'If the point was not one of the edge control point then...
            If CurrentPointDragged.x <> 1 And CurrentPointDragged.x <> G(Index).GridDiamension.x + 1 _
            And CurrentPointDragged.y <> 1 And CurrentPointDragged.y <> G(Index).GridDiamension.y + 1 Then
                Dim pt(1 To 6) As PointAPI, Region As Long
                'get all the 6 points connected to the dragged point
                pt(1) = G(Index).GridPoint(CurrentPointDragged.x - 1, CurrentPointDragged.y - 1)
                pt(2) = G(Index).GridPoint(CurrentPointDragged.x, CurrentPointDragged.y - 1)
                pt(3) = G(Index).GridPoint(CurrentPointDragged.x + 1, CurrentPointDragged.y)
                pt(4) = G(Index).GridPoint(CurrentPointDragged.x + 1, CurrentPointDragged.y + 1)
                pt(5) = G(Index).GridPoint(CurrentPointDragged.x, CurrentPointDragged.y + 1)
                pt(6) = G(Index).GridPoint(CurrentPointDragged.x - 1, CurrentPointDragged.y)
                'create a region from these point
                Region = CreatePolygonRgn(pt(1), 6, 1)
                
                'if the dragged control point is not within the region then _
                show the user a message box informing them, then restore the _
                control point to its original position.
                If PtInRegion(Region, x / ZoomValue, y / ZoomValue) = 0 Then
                    MsgBox "Triangles are overlapping!, you cannot overlap triangles.", vbInformation Or vbOKOnly, "Overlap detected"
                    G(Index).GridPoint(CurrentPointDragged.x, CurrentPointDragged.y) = LastPoint
                    DrawGrid P(Index), G(Index), ZoomValue
                End If
                'Clear the region from memory
                DeleteObject Region
            Else
                'if the points were the edge points then...
                
                'If it was along the left or the right of the grid then
                If CurrentPointDragged.x <> 1 And CurrentPointDragged.x <> G(Index).GridDiamension.x + 1 Then
                    If x > ZoomValue * G(Index).GridPoint(CurrentPointDragged.x + 1, CurrentPointDragged.y).x _
                    Or x < ZoomValue * G(Index).GridPoint(CurrentPointDragged.x - 1, CurrentPointDragged.y).x Then
                        MsgBox "Triangles are overlapping!, you cannot overlap triangles.", vbInformation Or vbOKOnly, "Overlap detected"
                        G(Index).GridPoint(CurrentPointDragged.x, CurrentPointDragged.y) = LastPoint
                        DrawGrid P(Index), G(Index), ZoomValue
                    End If
                End If
                
                'If it was along the top or the bottom of the grid then
                If CurrentPointDragged.y <> 1 And CurrentPointDragged.y <> G(Index).GridDiamension.y + 1 Then
                    If y > ZoomValue * G(Index).GridPoint(CurrentPointDragged.x, CurrentPointDragged.y + 1).y _
                    Or y < ZoomValue * G(Index).GridPoint(CurrentPointDragged.x, CurrentPointDragged.y - 1).y Then
                        MsgBox "Triangles are overlapping!, you cannot overlap triangles.", vbInformation Or vbOKOnly, "Overlap detected"
                        G(Index).GridPoint(CurrentPointDragged.x, CurrentPointDragged.y) = LastPoint
                        DrawGrid P(Index), G(Index), ZoomValue
                    End If
                End If
            End If
            
            SelIndex = IIf(Index = 1, 2, 1)
            'this will redraw the grid on the opposit picture box to the one selected
            'hence removing the indicator on the current control point
            P(SelIndex).Cls
            DrawGrid P(SelIndex), G(SelIndex), ZoomValue
            
            'Indicate that a control point is no longer being dragged
            CurrentPointDragged.x = -1
            CurrentPointDragged.y = -1
            'indicate that morph has been changed
            MorphDirty = True
        End If
    End If
End Sub

Private Sub VScroll1_Change()
    P(1).top = -VScroll1.Value
    P(2).top = -VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub
