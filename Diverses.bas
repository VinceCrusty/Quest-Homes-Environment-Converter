Attribute VB_Name = "Diverses"
Option Explicit

Public idr2 As String
Public BuildPath As String
Public answer As Boolean
Public start_pano As Boolean
Public pataud As String
Public u As String
Public aud As String

Public use_pic As String
Public rota(40) As String
Public gltf1 As String
Public gltf2 As String
Public sva As Integer
Public vp As Integer
Public renunp As Boolean
Public apppath As String
Public Mak As Boolean

'------------------------------------------------------------------

Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

Public Declare Function GetPrivateProfileString Lib "Kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "Kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long

'---------------------------------------------------

Private Const BUFFERSIZE As Long = 65535

Public Enum eImageType
    itUNKNOWN = 0
    itGIF = 1
    itJPEG = 2
    itPNG = 3
    itBMP = 4
End Enum

Public m_Width As Long
Public m_Height As Long
Public m_Depth As Byte
Public m_ImageType As eImageType

Public Function GetINISetting(ByVal sHeading As String, ByVal sKey As String, sINIFileName) As String
    
    Const cparmLen = 400
    
    Dim sReturn As String * cparmLen
    Dim sDefault As String * cparmLen
    Dim lLength As Long
    lLength = GetPrivateProfileString(sHeading, _
            sKey, _
            sDefault, _
            sReturn, _
            cparmLen, _
            sINIFileName)
            
    GetINISetting = Mid(sReturn, 1, lLength)
End Function

Public Function PutINISetting(ByVal sHeading As String, ByVal sKey As String, ByVal sSetting As String, sINIFileName) As Boolean
    
Const cparmLen = 400

Dim sReturn As String * cparmLen
Dim sDefault As String * cparmLen
Dim aLength As Long
aLength = WritePrivateProfileString(sHeading, sKey, sSetting, sINIFileName)

Form1.txtOutputs.Text = Form1.txtOutputs.Text & vbNewLine & vbNewLine & "Update Setting: " & sHeading & " > " & sKey & " = " & sSetting & vbNewLine
Form1.txtOutputs.SelStart = Len(Form1.txtOutputs.Text)

PutINISetting = True
    
End Function

Public Function HexIt(ByVal MyColour As Long) As String

Dim Reply As String
Reply = Hex(MyColour)
If Len(Reply) < 6 Then
Reply = String$(6 - Len(Reply), "0") + Reply
End If
HexIt = "&H00" + Mid$(Reply, 1, 2) + Mid$(Reply, 3, 2) + Mid$(Reply, 5, 2) + "&"

End Function

Public Function HTC(ByRef HexColor As String) As Long

If HexColor = "" Then Exit Function
HTC = Int(Left$(HexColor, Len(HexColor) - 1))
    
End Function

Public Sub Pause(Seconds As Single)

Dim Timera1 As Single, Timera2 As Single, currentDate As Date

currentDate = Date
Timera1 = Timer + Seconds
Timera2 = Timera1 - 86400
While ((Timer() < Timera1) And (currentDate = Date)) Or ((Timer() < Timera2) And (currentDate + 1 = Date))
  DoEvents
Wend

End Sub

Public Function ExtractFile(ByVal PathName As String) As String

On Error Resume Next

Dim f As String
Dim n As Integer

f$ = PathName
Do
    n% = InStr(f$, "\")
    If n% > 0 Then f$ = Right$(f$, Len(f$) - n%)
Loop While n% > 0
Do
    n% = InStr(f$, "/")
    If n% > 0 Then f$ = Right$(f$, Len(f$) - n%)
Loop While n% > 0

ExtractFile = f$

End Function

Public Function Reset()

ChDrive Left$(apppath, 2)
ChDir apppath

End Function
Public Function Message(mesa1 As String, Optional big As Boolean)

On Error Resume Next

If big = True Then
   Form3.Label1.FontSize = 14
Else
   Form3.Label1.FontSize = 18
End If
Form3.Label1.Caption = mesa1
Call Form3.Show(vbModal)

End Function

Public Function Question(mesa1 As String, Optional big As Boolean) As Boolean

On Error Resume Next

If big = True Then
   Form8.Label1.FontSize = 14
Else
   Form8.Label1.FontSize = 18
End If
Form8.Label1.Caption = mesa1
Call Form8.Show(vbModal)
Question = answer

End Function

Public Sub ReadImageInfo(sFilename As String)

On Error Resume Next

Dim bBuf(BUFFERSIZE) As Byte
Dim iFN As Integer

m_Width = 0
m_Height = 0
m_Depth = 0
m_ImageType = itUNKNOWN
iFN = FreeFile
Open sFilename For Binary As iFN
Get #iFN, 1, bBuf()
Close iFN
If bBuf(0) = 137 And bBuf(1) = 80 And bBuf(2) = 78 Then
    m_ImageType = itPNG
    Select Case bBuf(25)
        Case 0
            m_Depth = bBuf(24)
        Case 2
            m_Depth = bBuf(24) * 3
        Case 3
            m_Depth = 8
        Case 4
            m_Depth = bBuf(24) * 2
        Case 6
            m_Depth = bBuf(24) * 4
        Case Else
            m_ImageType = itUNKNOWN
    End Select
    If m_ImageType Then
        m_Width = Mult(bBuf(19), bBuf(18))
        m_Height = Mult(bBuf(23), bBuf(22))
    End If
End If
If bBuf(0) = 71 And bBuf(1) = 73 And bBuf(2) = 70 Then
    m_ImageType = itGIF
    m_Width = Mult(bBuf(6), bBuf(7))
    m_Height = Mult(bBuf(8), bBuf(9))
    m_Depth = (bBuf(10) And 7) + 1
End If
If bBuf(0) = 66 And bBuf(1) = 77 Then
    m_ImageType = itBMP
    m_Width = Mult(bBuf(18), bBuf(19))
    m_Height = Mult(bBuf(22), bBuf(23))
    m_Depth = bBuf(28)
End If
If m_ImageType = itUNKNOWN Then
    Dim lPos As Long
    Do
        If (bBuf(lPos) = &HFF And bBuf(lPos + 1) = &HD8 _
             And bBuf(lPos + 2) = &HFF) _
             Or (lPos >= BUFFERSIZE - 10) Then Exit Do
        lPos = lPos + 1
    Loop
    lPos = lPos + 2
    If lPos >= BUFFERSIZE - 10 Then Exit Sub
    Do
        Do
            If bBuf(lPos) = &HFF And bBuf(lPos + 1) _
           <> &HFF Then Exit Do
            lPos = lPos + 1
            If lPos >= BUFFERSIZE - 10 Then Exit Sub
        Loop
        lPos = lPos + 1
        Select Case bBuf(lPos)
            Case &HC0 To &HC3, &HC5 To &HC7, &HC9 To &HCB, _
            &HCD To &HCF
                Exit Do
        End Select
        lPos = lPos + Mult(bBuf(lPos + 2), bBuf(lPos + 1))
        If lPos >= BUFFERSIZE - 10 Then Exit Sub
    Loop
    m_ImageType = itJPEG
    m_Height = Mult(bBuf(lPos + 5), bBuf(lPos + 4))
    m_Width = Mult(bBuf(lPos + 7), bBuf(lPos + 6))
    m_Depth = bBuf(lPos + 8) * 8
End If
    
End Sub
Private Function Mult(lsb As Byte, msb As Byte) As Long

On Error Resume Next

Mult = lsb + (msb * CLng(256))
    
End Function

Public Function LoadPicture2(strPath As String, pic As PictureBox)

'LoadPicture "C:\Users\Ricco\Desktop\1\1.png", Picture1
With CreateObject("WIA.ImageFile")
    .LoadFile (strPath)
    pic.Picture = .FileData.Picture
End With
pic.ScaleMode = 3
pic.AutoRedraw = True
pic.PaintPicture pic.Picture, _
0, 0, pic.ScaleWidth, pic.ScaleHeight, _
0, 0, pic.Picture.Width / 26.46, _
pic.Picture.Height / 26.46
pic.Picture = pic.Image
    
End Function
