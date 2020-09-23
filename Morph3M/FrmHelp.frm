VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Morpher3"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmHelp.frx":0000
   ScaleHeight     =   145
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   269
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ok"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Click here to vote now!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Contact me before using any parts of this program."
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   4020
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "nirpaudyal@hotmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   720
      MouseIcon       =   "FrmHelp.frx":13A46
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1200
      Width           =   2325
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   705
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Niranjan Paudyal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   0
      Picture         =   "FrmHelp.frx":13D50
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   3975
   End
End
Attribute VB_Name = "FrmAbout"
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
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWND As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
    ShellExecute Me.HWND, "open", "www.planet-source-code.com", vbNull, vbNull, 1
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Label4_Click()
    ShellExecute Me.HWND, "open", "mailto:nirpaudyal@hotmail.com", vbNull, vbNull, 1
End Sub
