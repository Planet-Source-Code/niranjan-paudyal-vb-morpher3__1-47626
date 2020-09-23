VERSION 5.00
Begin VB.Form FrmNewMorph 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New morph"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5280
   Icon            =   "FrmNewMorph.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   352
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Load end image"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Load start image"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   2655
   End
   Begin VB.PictureBox B 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H000080FF&
      Height          =   2520
      Index           =   1
      Left            =   2640
      ScaleHeight     =   168
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   177
      TabIndex        =   5
      Top             =   0
      Width           =   2655
   End
   Begin VB.PictureBox B 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawMode        =   6  'Mask Pen Not
      DrawStyle       =   1  'Dash
      ForeColor       =   &H000080FF&
      Height          =   2520
      Index           =   0
      Left            =   0
      ScaleHeight     =   168
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   4
      Top             =   0
      Width           =   2625
   End
   Begin VB.PictureBox P 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   1440
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   3
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox P 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   600
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   2
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ok"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   0
      Picture         =   "FrmNewMorph.frx":000C
      Stretch         =   -1  'True
      Top             =   2895
      Width           =   5295
   End
End
Attribute VB_Name = "FrmNewMorph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Part of VBMorpher3 written by Niranjan paudyal
'No parts of this program may be copied or used
'without contacting me first at
'nirpaudyal@hotmail.com (see the about box)
'If you like to use the avi file making module, or
'would like any help on that module, plese contact
'the author (contact address on mAVIDecs module)
Dim FileNames(0 To 1) As String 'The file names for the two pictures

Private Sub Command1_Click()    'Ok button
    
    Static MorphNo As Long
    'Check to see files are open
    If FileNames(0) = "" Or FileNames(1) = "" Then
        MsgBox "Two pictures must be chosen to morph between!", vbExclamation, "Information"
        Exit Sub
    End If
    'Check width of pictures and height of _
    pictures to make sure they are the same
    If P(0).ScaleWidth <> P(1).ScaleWidth Or _
    P(0).ScaleHeight <> P(1).ScaleHeight Then
            
        frmResize.Load_Pic 0, FileNames(0)
        frmResize.Load_Pic 1, FileNames(1)
        frmResize.Show 1
        If frmResize.Canceled Then
            Exit Sub
        Else
            FileNames(0) = frmResize.File_Name(0)
            FileNames(1) = frmResize.File_Name(1)
        End If
    End If
    'Load the Pictures onto the new form
    MorphNo = MorphNo + 1
    Unload Me
    Dim F As New FrmMain
    F.Caption = "Morph" & MorphNo
    F.NewMorph FileNames(0), FileNames(1)
    F.mnuMorphEditorOption_Click
End Sub
Private Sub Command2_Click()    'Cancel button
    Unload Me
End Sub

Private Sub Command3_Click(Index As Integer)
    Dim filename As String, res As Long, wr As Single
    Dim OD As cFileDlg
    Set OD = New cFileDlg
    
    'set dialog options
    With OD
        .DlgTitle = IIf(Index = 0, "Open start picture", "Open end picture")
        .Filter = "Picture file (*.bmp;*.jpg;*.gif;*.dib;*.wmf)|*.bmp;*.jpg;*.gif;*.dib;*.wmf"
        .OwnerHwnd = Me.HWND
        .InitDirectory = CurDir
    End With
    'Cancel was not pressed
    If OD.VBGetOpenFileName(filename) <> False Then
        B(Index).Cls
        'Clear picture box
        P(Index).Picture = LoadPicture("")
        'Load piucture
        P(Index).Picture = LoadPicture(filename)
        
        'Now stretch the pictures in their ratios so that they correctly fit
        'onto the pictur box
        If P(Index).ScaleHeight < P(Index).ScaleWidth Then
            wr = B(Index).ScaleWidth / P(Index).ScaleWidth
            B(Index).PaintPicture P(Index).Image, 0, B(0).ScaleHeight / 2 - (wr * P(Index).ScaleHeight) / 2, B(Index).ScaleWidth, wr * P(Index).ScaleHeight
        Else
            wr = B(Index).ScaleHeight / P(Index).ScaleHeight
            B(Index).PaintPicture P(Index).Image, B(Index).ScaleWidth / 2 - wr * P(Index).ScaleWidth / 2, 0, wr * P(Index).ScaleWidth, B(Index).ScaleHeight
        End If
        B(Index).Refresh
        FileNames(Index) = filename
    End If
    'remove the dialogbox from memory
    Set OD = Nothing
End Sub

Private Sub Form_Load()
    'Set file names to nothing
    FileNames(0) = ""
    FileNames(1) = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    'Remove pictures from memory
    For i = 0 To 1
        P(i).Picture = LoadPicture("")
        B(i).Picture = LoadPicture("")
    Next i
End Sub
