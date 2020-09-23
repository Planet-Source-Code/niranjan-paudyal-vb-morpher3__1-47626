Attribute VB_Name = "ModDib"
Option Explicit
'Part of VBMorpher3 written by Niranjan paudyal
'No parts of this program may be copied or used
'without contacting me first at
'nirpaudyal@hotmail.com (see the about box)
'If you like to use the avi file making module, or
'would like any help on that module, plese contact
'the author (contact address on mAVIDecs module)

Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function CreateCompatibleBitmap& Lib "gdi32" (ByVal hdc&, ByVal nWidth&, ByVal nHeight&)

Type BITMAPINFOHEADER
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type


Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  pBits() As Byte
End Type

Public Sub MakePicMem(ByRef Picture As PictureBox, tBM As BITMAPINFO, CommandNumber As Integer)
'CommandNumber=1      = store picture in memory
'CommandNumber=2      = refresh picture box
'CommandNumber=3      = Clear memory only
If CommandNumber = 1 Then
    Dim Bm As BITMAP
    Dim hdcNew As Long, OldHand As Long, Pic As Long, Ret As Long
    
    Pic = Picture.Image
    GetObject Pic, Len(Bm), Bm
    hdcNew = CreateCompatibleDC(0)
    OldHand = SelectObject(hdcNew, Pic)
    
    With tBM.bmiHeader
        .biSize = 40
        .biWidth = Bm.bmWidth
        .biHeight = -Bm.bmHeight
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = 0
    End With
  
    ReDim tBM.pBits(3, Bm.bmWidth - 1, Bm.bmHeight - 1) As Byte
    Ret = GetDIBits(hdcNew, Pic, 0, Bm.bmHeight, tBM.pBits(0, 0, 0), tBM, 0)
    
    DeleteObject OldHand
    DeleteDC hdcNew
End If

If CommandNumber = 2 Then SetDIBits Picture.hdc, Picture.Image, 0, -tBM.bmiHeader.biHeight, tBM.pBits(0, 0, 0), tBM, 0
If CommandNumber = 3 Then ReDim tBM.pBits(0, 0, 0) As Byte
End Sub
