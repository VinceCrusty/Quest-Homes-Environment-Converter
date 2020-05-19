Attribute VB_Name = "SaveImage"

Option Explicit
'///////////////////////////////////////////////////
'// SaveImageAs - Save hDC to Bitmap or Jpeg file //
'// Ed Wilk/Edgemeal - last updated Feb.06,2010   //
'///////////////////////////////////////////////////

'calling SaveImageAs:
'SaveImageAs "c:\mypicture.jpg", Picture1.hdc, Picture1.Width / Screen.TwipsPerPixelX, Picture1.Height / Screen.TwipsPerPixelY, 75)

Private Const BI_RGB As Long = 0
Private Const DIB_RGB_COLORS As Long = 0

Private Type BitmapFileHeader
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Private Type BitmapInfoHeader
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biDataSize As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Private Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgbReserved As Byte
End Type

Private Type BITMAPINFO
  bmiHeader As BitmapInfoHeader
  bmiColors As RGBQUAD
End Type

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' gdi+
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type

Private Type GdiplusStartupInput
   GdiplusVersion As Long
   DebugEventCallback As Long
   SuppressBackgroundThread As Long
   SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
   GUID As GUID
   NumberOfValues As Long
   Type As Long
   Value As Long
End Type

Private Type EncoderParameters
   Count As Long
   Parameter As EncoderParameter
End Type

Private Declare Function GdiplusStartup Lib "GDIPlus" ( _
   token As Long, _
   inputbuf As GdiplusStartupInput, _
   Optional ByVal outputbuf As Long = 0) As Long

Private Declare Function GdiplusShutdown Lib "GDIPlus" ( _
   ByVal token As Long) As Long

Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" ( _
   ByVal hbm As Long, _
   ByVal hPal As Long, _
   Bitmap As Long) As Long

Private Declare Function GdipDisposeImage Lib "GDIPlus" ( _
   ByVal Image As Long) As Long

Private Declare Function GdipSaveImageToFile Lib "GDIPlus" ( _
   ByVal Image As Long, _
   ByVal filename As Long, _
   clsidEncoder As GUID, _
   encoderParams As Any) As Long

Private Declare Function CLSIDFromString Lib "ole32" ( _
   ByVal str As Long, _
   id As GUID) As Long




Public Sub SaveImageAs(ByVal sFileName As String, ByVal Source_hDC As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal Quality As Long = 80)
    Dim sFileExt As String
    Dim myDIB As Long, myDC As Long, fNum As Long
    Dim bi24BitInfo As BITMAPINFO
    Dim fileheader As BitmapFileHeader
    Dim bitmapData() As Byte
    ' gdi
    Dim tSI As GdiplusStartupInput
    Dim lRes As Long, lGDIP As Long, lBitmap As Long
    Dim tJpgEncoder As GUID
    Dim tParams As EncoderParameters
    
    ' source hDC to DIB
    With bi24BitInfo.bmiHeader
      .biBitCount = 24
      .biCompression = BI_RGB
      .biPlanes = 1
      .biSize = Len(bi24BitInfo.bmiHeader)
      .biWidth = Width
      .biHeight = Height
      .biDataSize = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
      ReDim bitmapData(0 To .biDataSize - 1)
    End With
    myDC = CreateCompatibleDC(0)
    myDIB = CreateDIBSection(myDC, bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    SelectObject myDC, myDIB
    BitBlt myDC, 0, 0, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, Source_hDC, 0, 0, vbSrcCopy
    Call GetDIBits(myDC, myDIB, 0, bi24BitInfo.bmiHeader.biHeight, bitmapData(0), bi24BitInfo, DIB_RGB_COLORS)
        
    ' get file extension of filename to save as lower case.
    sFileExt = LCase$(GetFileExt(sFileName))
    ' Save image to file....
    Select Case sFileExt
        Case ".bmp"   ' save as bmp....
            With fileheader
                .bfType = &H4D42
                .bfOffBits = Len(fileheader) + Len(bi24BitInfo.bmiHeader)
                .bfSize = bi24BitInfo.bmiHeader.biDataSize + .bfOffBits
            End With
            fNum = FreeFile
            On Error GoTo BadFileName
            Open sFileName For Output As fNum
            Close fNum
            Open sFileName For Binary As fNum
            Put fNum, , fileheader
            Put fNum, , bi24BitInfo.bmiHeader
            Put fNum, , bitmapData()
            Close fNum
        Case ".jpg", ".png"
            tSI.GdiplusVersion = 1 ' Initialize GDI+
            lRes = GdiplusStartup(lGDIP, tSI)
            If lRes = 0 Then
                lRes = GdipCreateBitmapFromHBITMAP(myDIB, 0, lBitmap) ' Create the GDI+ bitmap from the image handle
                If lRes = 0 Then
                    If sFileExt = ".jpg" Then ' JPG
                        CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
                        ' Initialize the encoder parameters
                        tParams.Count = 1
                        With tParams.Parameter ' jpeg Quality
                          CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
                          .NumberOfValues = 1
                          .Type = 4
                          .Value = VarPtr(Quality)
                        End With
                    ElseIf sFileExt = ".png" Then ' PNG
                        CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
                    End If
                    ' Save the image
                    lRes = GdipSaveImageToFile(lBitmap, StrPtr(sFileName), tJpgEncoder, tParams)
                    ' Destroy the bitmap
                    GdipDisposeImage lBitmap
                End If
              ' Shutdown GDI+
              GdiplusShutdown lGDIP
            End If
            If lRes Then
                Err.Raise 5, , "Can not save image(GDI+ Error).:" & lRes
            End If
    End Select
Fini:
    DeleteObject myDIB
    DeleteDC myDC
    Exit Sub

BadFileName:
    Close fNum
    Err.Raise 5, , "Can not save BMP image.:" & lRes
    Resume Fini
End Sub
Private Function GetFileExt(sFile As String) As String
    ' example" returns ".exe"
    Dim i As Integer
    i = InStrRev(sFile, ".")
    If i Then
        GetFileExt = Mid$(sFile, i)
    Else
        GetFileExt = sFile ' if not found then just return the whole string
    End If
End Function



