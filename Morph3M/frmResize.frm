VERSION 5.00
Begin VB.Form frmResize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pictures not of same size"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frmResize.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   354
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox SavePic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1200
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   9
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      Width           =   1095
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
      TabIndex        =   7
      Top             =   990
      Width           =   2625
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
      TabIndex        =   6
      Top             =   990
      Width           =   2655
   End
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
      Height          =   735
      Index           =   1
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   2655
   End
   Begin VB.PictureBox tmpPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
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
      Height          =   735
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   0
      Picture         =   "frmResize.frx":000C
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please select which picture you wish to resize."
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "You have the option of resizing the pictures so they are of the same diamension."
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The two pictures are not of the same size."
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmResize"
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

Dim Has_canceled As Boolean
Dim iFileName(1) As String
Dim iFileDiamension(1) As PointAPI

Public Property Get File_Name(Index As Integer) As String
    File_Name = iFileName(Index)
End Property
Public Property Get Canceled() As Boolean
    Canceled = Has_canceled
End Property
Public Function Load_Pic(Index As Integer, filename As String)
    Dim H As Long, W As Long
    tmpPic.Picture = LoadPicture(filename)
    SetStretchBltMode B(Index).hdc, COLORONCOLOR
    If tmpPic.ScaleWidth < tmpPic.ScaleHeight Then
        H = B(Index).ScaleHeight
        W = (H / tmpPic.ScaleHeight) * tmpPic.ScaleWidth
        StretchBlt B(Index).hdc, B(Index).Width / 2 - W / 2, 0, W, H, tmpPic.hdc, 0, 0, tmpPic.ScaleWidth, tmpPic.ScaleHeight, vbSrcCopy
    Else
        W = B(Index).ScaleWidth
        H = (W / tmpPic.ScaleWidth) * tmpPic.ScaleHeight
        StretchBlt B(Index).hdc, 0, B(Index).Height / 2 - H / 2, W, H, tmpPic.hdc, 0, 0, tmpPic.ScaleWidth, tmpPic.ScaleHeight, vbSrcCopy
    End If
    Command3(Index).Caption = IIf(Len(filename) > 20, "..." & right(filename, 17), filename) & vbNewLine & "Width=" & B(Index).ScaleWidth & vbNewLine & "Height=" & B(Index).ScaleHeight
    iFileName(Index) = filename
    iFileDiamension(Index).x = tmpPic.Width
    iFileDiamension(Index).y = tmpPic.Height
End Function

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click(Index As Integer)
    Dim fFileName As String
    Dim TempPath As String
    Dim TmpLength As Long
    Dim OtherIndex As Integer
    Dim i As Long
    
    OtherIndex = IIf(Index = 1, 0, 1)
    For i = Len(iFileName(Index)) To 1 Step -1
        If Mid(iFileName(Index), i, 1) <> "\" Then
            fFileName = Mid(iFileName(Index), i, 1) & fFileName
        Else
            Exit For
        End If
    Next i
    
    fFileName = Mid(iFileName(Index), 1, Len(iFileName(Index)) - Len(fFileName)) & Mid(fFileName, 1, Len(fFileName) - 4) & "_resized.bmp"
    
    
    'TmpLength = 1
    'TmpLength = GetTempPath(TmpLength, TempPath)
    'TempPath = Space(TmpLength - 1)
    'GetTempPath TmpLength, TempPath
    'fFileName = TempPath & fFileName
    

    tmpPic.Picture = LoadPicture(iFileName(Index))
    SavePic.Width = iFileDiamension(OtherIndex).x
    SavePic.Height = iFileDiamension(OtherIndex).y
    
    SetStretchBltMode SavePic.hdc, HALFTONE
    StretchBlt SavePic.hdc, 0, 0, SavePic.Width, SavePic.Height, tmpPic.hdc, 0, 0, tmpPic.Width, tmpPic.Height, vbSrcCopy
    
    If Dir(fFileName) <> "" Then Kill (fFileName)
    
    SavePicture SavePic.Image, fFileName
    iFileName(Index) = fFileName
    Has_canceled = False
    Unload Me
    
End Sub

Private Sub Form_Load()
    Has_canceled = True
End Sub
