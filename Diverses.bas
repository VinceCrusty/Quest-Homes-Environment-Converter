Attribute VB_Name = "Diverses"
Option Explicit

Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long

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

Dim Timer1 As Single, Timer2 As Single, currentDate As Date

currentDate = Date
Timer1 = Timer + Seconds
Timer2 = Timer1 - 86400
While ((Timer() < Timer1) And (currentDate = Date)) Or ((Timer() < Timer2) And (currentDate + 1 = Date))
  DoEvents
Wend

End Sub
