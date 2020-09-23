VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type RGBTRIPLE
     rgbRed As Byte
     rgbGreen As Byte
     rgbBlue As Byte
End Type

Private Type RGBQUAD
     rgbBlue As Byte
     rgbGreen As Byte
     rgbRed As Byte
     rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER '40 bytes
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

Private Type BITMAPINFO256
     bmiHeader As BITMAPINFOHEADER
     bmiColors(0 To 255) As RGBQUAD
End Type

Private Type BITMAP '14 bytes
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Private Const BI_RGB = 0&

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO256, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBSection256 Lib "gdi32" Alias "CreateDIBSection" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO256, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Const DIB_RGB_COLORS = 0

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
'============================GIF STAFF================

Private Type GifScreenDescriptor
     logical_screen_width As Integer
     logical_screen_height As Integer
     Flags As Byte
     background_color_index As Byte
     pixel_aspect_ratio As Byte
End Type

Private Type GifImageDescriptor
     Left As Integer
     Top As Integer
     Width As Integer
     Height As Integer
     Format As Byte 'ImageFormat
End Type
'========Added by Wolfgang Goetz for transparent GIFs=====
Private Type CONTROLBLOCK '(April 8., 2002 --> Wolfgang Goetz)
    Blocksize As Byte
    Flags As Byte
    Delay As Integer
    TransParent_Color As Byte
    Terminator As Byte
End Type
Private Const GIF89a = "GIF89a"
Private Const CtrlIntro As Byte = &H21
Private Const CtrlLabel As Byte = &HF9
'========================================================
Const GIF87a = "GIF87a"

Const GifTerminator As Byte = &H3B
Const ImageSeparator As Byte = &H2C
Const CHAR_BIT = 8
Const CodeSize As Byte = 9
Const ClearCode = 256
Const EndCode  As Integer = 257
Const FirstCode = 258
Const LastCode As Integer = 511
Const MAX_CODE = LastCode - FirstCode

Private colTable As New Collection
Private fn As Integer
Private gifPalette(0 To 255) As RGBTRIPLE
Private bit_position As Integer
Private code_count As Integer
Private data_buffer(255) As Byte
Private aPower2(31) As Long
Private picWidth As Long, picHeight As Long
Private IsBusy As Boolean
'Private cDibSave As New cDIBSectionSave
Private m_dib As New cDIBSection
Public Event Progress(ByVal Percents As Integer)
Dim bi As BITMAPINFO256
Dim bm As BITMAP

Public Function SaveGIF(ByVal pic As StdPicture, ByVal sFileName As String, _
       Optional hdc As Long = 0, Optional UseTrans As Boolean = False, _
       Optional ByVal TransColor As Long = 0) As Boolean
   If IsBusy Then Exit Function
   Dim scr As GifScreenDescriptor, im As GifImageDescriptor
   Dim hDCScn As Long, OldObj As Long, Src_hDc As Long
   Dim hDib256 As Long, hDC256 As Long, OldObj256 As Long
   Dim buf() As Byte, data As Byte, TransIndex As Byte
   Dim i As Long, j As Long, clr As Long
   Dim bFound As Boolean
   Dim intCode As Integer, nCount  As Integer
   Dim sPrefix As String, sByte As String
   Dim tempPic As StdPicture
   IsBusy = True

   m_dib.CreateFromPicture pic
   Call GetObjectAPI(pic, Len(bm), bm)
   picWidth = bm.bmWidth
   picHeight = bm.bmHeight
   ReDim buf(CLng(((picWidth + 3) \ 4) * 4), picHeight) As Byte

'Prepare DC for paintings
   hDCScn = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   hDC256 = CreateCompatibleDC(hDCScn)
   If hdc = 0 Then
      Src_hDc = CreateCompatibleDC(hDCScn)
      OldObj = SelectObject(Src_hDc, pic)
   Else
      Src_hDc = hdc
   End If
   DeleteDC hDCScn
   
   hDib256 = CreateDib256(hdc, bi)
   
   If hDib256 <> 0 Then
      OldObj256 = SelectObject(hDC256, hDib256)
      Call BitBlt(hDC256, 0, 0, picWidth, picHeight, Src_hDc, 0, 0, vbSrcCopy)
      For i = 0 To picHeight - 1
          Call GetDIBits(hDC256, hDib256, i, 1, buf(0, picHeight - i), bi, 0)
      Next
   Else
      With bi.bmiHeader
          .biSize = Len(bi.bmiHeader)
          .biWidth = picWidth
          .biHeight = picHeight
          .biPlanes = 1
          .biBitCount = 8
          .biCompression = BI_RGB
      End With
      For i = 0 To picHeight - 1
          Call GetDIBits(Src_hDc, pic, i, 1, buf(0, picHeight - i), bi, DIB_RGB_COLORS)
      Next
   End If
'Fill gif file info
   For i = 0 To 255
       gifPalette(i).rgbBlue = bi.bmiColors(i).rgbBlue
       gifPalette(i).rgbGreen = bi.bmiColors(i).rgbGreen
       gifPalette(i).rgbRed = bi.bmiColors(i).rgbRed
       If Not bFound Then
          clr = RGB(gifPalette(i).rgbRed, gifPalette(i).rgbGreen, gifPalette(i).rgbBlue)
          If clr = TransColor Then
             TransIndex = i: bFound = True
          End If
       End If
   Next
'  If TransColor = 0 Then TransIndex = 0
   scr.background_color_index = 0
   scr.Flags = &HF7 '256-color gif with global color map
   scr.pixel_aspect_ratio = 0
   
   im.Format = &H7 'GlobalNonInterlaced
   im.Height = picHeight
   im.Width = picWidth
  
   If FileExists(sFileName) Then Kill sFileName
  
   fn = FreeFile
   Open sFileName For Binary As fn
'Write GIF header and header info
     If UseTrans = True Then '(April 8., 2002 --> Wolfgang Goetz)
        Put #fn, , GIF89a
     Else
        Put #fn, , GIF87a
     End If
     Put #fn, , scr
     Put #fn, , gifPalette
 '(April 8., 2002 --> Wolfgang Goetz)
     If UseTrans = True Then
        Put #fn, , CtrlIntro
        Put #fn, , CtrlLabel
        Dim cb As CONTROLBLOCK
        cb.Blocksize = 4 'Always 4
        cb.Flags = 9 'Packed = 00001001 (If Bit 0 = 1: Use Transparency)
        cb.Delay = 0
        cb.TransParent_Color = TransIndex
        cb.Terminator = 0 'Always 0
        Put #fn, , cb
     End If
     Put #fn, , ImageSeparator
     Put #fn, , im
     data = CodeSize - 1
     Put #fn, , data
     data_buffer(0) = 0
     bit_position = CHAR_BIT
'Process pixels data using LZW/GIF compression
     For i = 1 To picHeight
         Reinitialize
         sPrefix = ""
         intCode = buf(0, i)
         On Error Resume Next
         For j = 1 To picWidth - 1
             sByte = MyFormat(buf(j, i))
             sPrefix = sPrefix & sByte
             intCode = colTable(sPrefix)
             If Err <> 0 Then 'Prefix wasn't in collection - save it and output code
                nCount = colTable.Count
                If nCount = MAX_CODE Then Reinitialize
                 colTable.Add nCount + FirstCode, sPrefix
                 OutputBits intCode, CodeSize
                 sPrefix = sByte
                 intCode = buf(j, i)
                 Err.Clear
             End If
         Next
         OutputBits intCode, CodeSize
         If i Mod 10 = 0 Then
            RaiseEvent Progress(i * 100 / picHeight)
            DoEvents
         End If
     Next
     OutputCode (EndCode)
     For i = 0 To data_buffer(0)
         Put #fn, , data_buffer(i)
     Next
     data = 0
     Put #fn, , data
     Put #fn, , GifTerminator
   Close fn
   Erase buf
   If hdc = 0 Then
      SelectObject Src_hDc, OldObj
      DeleteDC Src_hDc
   End If
   SelectObject hDC256, OldObj256
   DeleteObject hDib256
   DeleteDC hDC256
   SaveGIF = True
   IsBusy = False
End Function

Private Sub OutputBits(Value As Integer, Count As Integer)
   Dim i As Integer, bit As Integer
   Do While i < Count
      If bit_position = CHAR_BIT Then
         If data_buffer(0) = 255 Then
            Put #fn, , data_buffer
            data_buffer(0) = 1
         Else
            data_buffer(0) = data_buffer(0) + 1
         End If
         data_buffer(data_buffer(0)) = 0
         bit_position = 0
       End If
       bit = Sgn(Power2(i) And Value)
       If bit > 0 Then data_buffer(data_buffer(0)) = Power2(bit_position) Or data_buffer(data_buffer(0))
       i = i + 1: bit_position = bit_position + 1
   Loop
End Sub

Private Sub OutputCode(code As Integer)
    code_count = code_count + 1
    If code_count > LastCode Then
      code_count = FirstCode
      Call OutputBits(ClearCode, CodeSize)
      ClearTable
    End If
    Call OutputBits(code, CodeSize)
End Sub

Private Sub ClearTable()
   Set colTable = Nothing
   Set colTable = New Collection
End Sub

Private Sub Reinitialize()
   ClearTable
   Call OutputBits(ClearCode, CodeSize)
End Sub

Private Function FileExists(ByVal strPathName As String) As Boolean
   Dim af As Long
   af = GetFileAttributes(strPathName)
   FileExists = (af <> -1)
End Function

Private Function Power2(ByVal i As Integer) As Long
    If aPower2(0) = 0 Then
       aPower2(0) = &H1&
       aPower2(1) = &H2&
       aPower2(2) = &H4&
       aPower2(3) = &H8&
       aPower2(4) = &H10&
       aPower2(5) = &H20&
       aPower2(6) = &H40&
       aPower2(7) = &H80&
       aPower2(8) = &H100&
       aPower2(9) = &H200&
       aPower2(10) = &H400&
       aPower2(11) = &H800&
       aPower2(12) = &H1000&
       aPower2(13) = &H2000&
       aPower2(14) = &H4000&
       aPower2(15) = &H8000&
       aPower2(16) = &H10000
       aPower2(17) = &H20000
       aPower2(18) = &H40000
       aPower2(19) = &H80000
       aPower2(20) = &H100000
       aPower2(21) = &H200000
       aPower2(22) = &H400000
       aPower2(23) = &H800000
       aPower2(24) = &H1000000
       aPower2(25) = &H2000000
       aPower2(26) = &H4000000
       aPower2(27) = &H8000000
       aPower2(28) = &H10000000
       aPower2(29) = &H20000000
       aPower2(30) = &H40000000
       aPower2(31) = &H80000000
    End If
    Power2 = aPower2(i)
End Function

Private Function MyFormat(ByVal s As String) As String
   MyFormat = Right$("00" & s, 3)
End Function

Private Function CreateDib256(ByVal h_Dc As Long, bi As BITMAPINFO256) As Long
Dim lScanSize As Long
Dim lptr As Long, lIndex As Long
Dim r As Long, g As Long, b As Long
Dim rA As Long, gA As Long, bA As Long
   
Dim tBI256 As BITMAPINFO256
Dim hDib As Long, bSuccess As Boolean
Dim lHDC As Long, hBmpOld As Long, lHDCWOrk As Long, i As Long
Dim cDIBWork As cDIBSection
Dim bBltIn As Boolean
Dim cPal As New cPalette

   With bi.bmiHeader
       .biSize = Len(bi.bmiHeader)
       .biWidth = picWidth
       .biHeight = picHeight
       .biPlanes = 1
       .biBitCount = 8
       .biCompression = BI_RGB
       lScanSize = (picWidth + picWidth Mod 4)
       .biSizeImage = lScanSize * picHeight
   End With
   
    Set cPal = New cPalette
    Set cDIBWork = New cDIBSection
    cDIBWork.Create m_dib.Width, m_dib.Height
    If Form1.Option1.Value = True Then
        cPal.CreateOptimal m_dib
    Else
        cPal.CreateWebSafe
    End If
    Call pbCreate256ColourDIBSection(hDib, bi, lptr, m_dib.Width, m_dib.Height, cPal)
    CreateDib256 = CreateDIBSection256(h_Dc, bi, DIB_RGB_COLORS, lptr, 0, 0)
   
End Function

Private Function pbCreate256ColourDIBSection(ByRef hDib As Long, ByRef tBI As BITMAPINFO256, ByRef lptr As Long, ByVal lWidth As Long, ByVal lheight As Long, Optional ByRef cP As cPalette = Nothing) As Long
Dim lScanSize As Long
Dim lHDC As Long
Dim i As Long
Dim iMax As Long

   lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   With tBI.bmiHeader
       .biSize = Len(tBI.bmiHeader)
       .biWidth = lWidth
       .biHeight = lheight
       .biPlanes = 1
       .biBitCount = 8
       .biCompression = BI_RGB
       lScanSize = (lWidth + lWidth Mod 4)
       .biSizeImage = lScanSize * .biHeight
   End With
   iMax = 255
   If iMax >= cP.Count Then
      iMax = cP.Count - 1
   End If
   For i = 0 To iMax
      With tBI.bmiColors(i)
         .rgbBlue = cP.Blue(i + 1)
         .rgbGreen = cP.Green(i + 1)
         .rgbRed = cP.Red(i + 1)
      End With
   Next
'   bi = tBI
'   hDib = CreateDIBSection256( _
'           lHDC, _
'           tBI, _
'           DIB_RGB_COLORS, _
'           lptr, _
'           0, 0)
   pbCreate256ColourDIBSection = hDib
   DeleteDC lHDC
End Function
