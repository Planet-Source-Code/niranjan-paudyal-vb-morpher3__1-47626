VERSION 5.00
Begin VB.MDIForm MDIFormMain 
   BackColor       =   &H8000000C&
   Caption         =   "VB Morpher3"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4620
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewMorph 
         Caption         =   "New Morph"
      End
      Begin VB.Menu mnuOpenMorph 
         Caption         =   "Open Morph"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "MDIFormMain"
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
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal HWND As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Sub MDIForm_Load()
    Dim lhWNd As Long, lHDC As Long
    'this will check to see if the system BPP if
    'above 24 BPP
    lhWNd = GetDesktopWindow()
    lHDC = GetDC(lhWNd)
    If GetDeviceCaps(lHDC, 12) < 24 Then
        MsgBox "This program can only run in a color depth of 24 bits or above.", vbExclamation, "Invalid color depth"
        Unload Me
    End If
End Sub
Private Sub mnuAbout_Click()
    FrmAbout.Show 1
End Sub

Public Sub mnuNewMorph_Click()
    FrmNewMorph.Show 1
End Sub
Public Sub mnuOpenMorph_Click()
    Dim filename As String, res As Long
    Dim OD As cFileDlg
    Set OD = New cFileDlg
    'set dialog options
    With OD
        .DlgTitle = "Open morph file"
        .Filter = "Morph file (*.mrf)|*.mrf"
        .OwnerHwnd = Me.HWND
        .InitDirectory = CurDir
        '.Flags =
    End With
    
    'Cancel was not pressed
    If OD.VBGetOpenFileName(filename) <> False Then
        Dim F As New FrmMain
        F.OpenMorph filename
    End If
    'remove the dialogbox from memory
    Set OD = Nothing
End Sub
