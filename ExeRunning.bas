Attribute VB_Name = "ExeRunning"
Option Explicit

Private Declare Function CreateToolhelpSnapshot Lib "Kernel32" _
  Alias "CreateToolhelp32Snapshot" ( _
  ByVal lFlgas As Long, _
  ByVal lProcessID As Long) As Long
 
Private Declare Function ProcessFirst Lib "Kernel32" _
  Alias "Process32First" ( _
  ByVal hSnapshot As Long, _
  uProcess As PROCESSENTRY32) As Long
 
Private Declare Function ProcessNext Lib "Kernel32" _
  Alias "Process32Next" ( _
  ByVal hSnapshot As Long, _
  uProcess As PROCESSENTRY32) As Long
 
Private Declare Sub CloseHandle Lib "Kernel32" ( _
  ByVal hPass As Long)
 
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const MAX_PATH As Long = 260
 
Private Type PROCESSENTRY32
  dwSize As Long
  cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As Long
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwflags As Long
  szexeFile As String * MAX_PATH
End Type

' Prüft, ob eine EXE-Datei bereits ausgeführt wird
Public Function IsEXERunning(ByVal sFilename As String) As Long
 
  Dim lSnapshot As Long
  Dim uProcess As PROCESSENTRY32
  Dim nResult As Long
 
  ' "Snapshot" des aktuellen Prozess ermitteln
  lSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
  If lSnapshot <> 0 Then
    uProcess.dwSize = Len(uProcess)
 
    ' Ersten Prozess ermitteln
    nResult = ProcessFirst(lSnapshot, uProcess)
 
    Do Until nResult = 0
      ' Prozessliste durchlaufen
      If InStr(LCase$(uProcess.szexeFile), LCase$(sFilename)) > 0 Then
        ' Jepp - EXE gefunden
        IsEXERunning = True
        Exit Do
      End If
 
      ' nächster Prozess
      nResult = ProcessNext(lSnapshot, uProcess)
    Loop
 
    ' Handle schliessen
    CloseHandle lSnapshot
  End If
End Function

