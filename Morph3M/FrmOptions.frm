VERSION 5.00
Begin VB.Form FrmOptions 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "FrmOptions.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   241
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   281
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ok"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   209
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   281
      TabIndex        =   2
      Top             =   0
      Width           =   4215
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   16
         Text            =   "10"
         ToolTipText     =   "Radius of the control point"
         Top             =   2880
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save as AVI file"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2760
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save as BMP images"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2760
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   13
         Text            =   "10"
         ToolTipText     =   "Radius of the control point"
         Top             =   2520
         Width           =   615
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F7F7F7&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   60
         Left            =   0
         ScaleHeight     =   4
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   281
         TabIndex        =   11
         Top             =   255
         Width           =   4215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3600
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "10"
         ToolTipText     =   "Number of vertical sections on grid"
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "10"
         ToolTipText     =   "The total number of frames on the morph"
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "10"
         ToolTipText     =   "Number of horizontal sections on grid"
         Top             =   2280
         Width           =   975
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1890
         Left            =   -15
         TabIndex        =   4
         ToolTipText     =   "Location to save the morphed images."
         Top             =   285
         Width           =   4245
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -30
         TabIndex        =   3
         ToolTipText     =   "Drive to save the morphed images."
         Top             =   -30
         Width           =   4275
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FPS"
         Height          =   255
         Left            =   3360
         TabIndex        =   17
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Control point radius"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   12
         Top             =   2520
         Width           =   1410
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         Index           =   0
         X1              =   0
         X2              =   281
         Y1              =   145
         Y2              =   145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Grid width"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Grid height"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   6
         Top             =   2280
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total frames"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   900
      End
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   0
      Picture         =   "FrmOptions.frx":000C
      Stretch         =   -1  'True
      Top             =   3150
      Width           =   4215
   End
End
Attribute VB_Name = "FrmOptions"
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

Public GridX As Long, GridY As Long, NoFrames As Long, _
CPRadius As Long, MorphPath As String, _
SaveBMPformat As Boolean, FPS As Long

Public Cancel As Boolean    'Did the user press the cancel button?
Dim LastDrive As String 'Last successful drive access

Private Sub Command1_Click()    'Cancel key
    Unload Me
End Sub

Private Sub Command2_Click()    'OK key
    'assign all variabls
    Cancel = False
    GridX = Text2
    GridY = Text3
    NoFrames = Text1 - 1
    CPRadius = Text4
    MorphPath = Dir1
    SaveBMPformat = Option1(0).Value
    FPS = Text5
    Unload Me
End Sub

Private Sub Drive1_Change()
    'If the drive was not found or there was any error the show error
    On Error GoTo Er
    'otherwise
    Dir1 = Drive1
    'get the last successful drive
    LastDrive = Drive1
    Exit Sub
Er:
    MsgBox "Drive not ready!", vbExclamation, "Error"
    'set new drive as the last successful drive
    Drive1 = LastDrive
End Sub

Private Sub Form_Load()
    'by default,set cancel to true, this is because the user might
    'perss the X button on the title bar
    Cancel = True
    LastDrive = Drive1
End Sub
Private Sub Option1_Click(Index As Integer)
    'if the first index is selected then allow the
    'user to enter the FPS value for the avi file.
    Text5.Enabled = (Index = 1)
    Label5.Enabled = (Index = 1)
End Sub

'The rest of this module deals with data validation.
'Checks to see if values entered in the text boxes are
'within bounds.
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim k As Integer
    k = KeyAscii
    If (k < 48 Or k > 57) And (k <> 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    With Text1
        'Total framse must be >= 2
        If (.Text = "0" Or .Text = "" Or Val(.Text) < 2) Then
            .Text = "2"
            .SelLength = 4
        End If
    End With
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim k As Integer

    k = KeyAscii
    If (k < 48 Or k > 57) And (k <> 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    With Text2
        If (.Text = "0" Or .Text = "" Or Val(.Text) < 0) Then
            .Text = "1"
            .SelLength = 4
        End If
    End With
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    Dim k As Integer

    k = KeyAscii
    If (k < 48 Or k > 57) And (k <> 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
    With Text3
        If (.Text = "0" Or .Text = "" Or Val(.Text) < 0) Then
            .Text = "1"
            .SelLength = 4
        End If
    End With
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    Dim k As Integer

    k = KeyAscii
    If (k < 48 Or k > 57) And (k <> 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
    With Text4
        'Control point radius must be between 2-15
        If (.Text = "0" Or .Text = "" Or Val(.Text) < 2) Then
            .Text = "2"
            .SelLength = 4
        End If
        If Val(.Text) > "15" Then
            .Text = "15"
            .SelLength = 4
        End If
    End With
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
    Dim k As Integer
    k = KeyAscii
    If (k < 48 Or k > 57) And (k <> 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
    With Text5
        'FPS must be between 1-30
        If (.Text = "0" Or .Text = "" Or Val(.Text) < 1) Then
            .Text = "1"
            .SelLength = 4
        End If
        If Val(.Text) > "30" Then
            .Text = "30"
            .SelLength = 4
        End If
    End With
End Sub
