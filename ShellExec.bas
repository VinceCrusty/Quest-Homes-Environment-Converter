Attribute VB_Name = "ShellExec"
Option Explicit

Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function GetNamedPipeInfo Lib "kernel32" (ByVal hNamedPipe As Long, lType As Long, lLenOutBuf As Long, lLenInBuf As Long, lMaxInstances As Long) As Long
   
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As Any, lpProcessInformation As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


'Purpose     :  Synchronously runs a DOS command line and returns the captured screen output.
'Inputs      :  sCommandLine                The DOS command line to run.
'               [bShowWindow]               If True displays the DOS output window.
'Outputs     :  Returns the screen output
'Notes       :  This routine will work only with those program that send their output to
'               the standard output device (stdout).
'               Windows NT ONLY.
'Revisions   :

Public Function ShellExecuteCapture(sCommandLine As String, Optional bShowWindow As Boolean = False) As String
    Const clReadBytes As Long = 256, INFINITE As Long = &HFFFFFFFF
    Const STARTF_USESHOWWINDOW = &H1, STARTF_USESTDHANDLES = &H100&
    Const SW_HIDE = 0, SW_NORMAL = 1
    Const NORMAL_PRIORITY_CLASS = &H20&
    
    Const PIPE_CLIENT_END = &H0     'The handle refers to the client end of a named pipe instance. This is the default.
    Const PIPE_SERVER_END = &H1     'The handle refers to the server end of a named pipe instance. If this value is not specified, the handle refers to the client end of a named pipe instance.
    Const PIPE_TYPE_BYTE = &H0      'The named pipe is a byte pipe. This is the default.
    Const PIPE_TYPE_MESSAGE = &H4   'The named pipe is a message pipe. If this value is not specified, the pipe is a byte pipe
    
    
    Dim tProcInfo As PROCESS_INFORMATION, lRetVal As Long, lSuccess As Long
    Dim tStartupInf As STARTUPINFO
    Dim tSecurAttrib As SECURITY_ATTRIBUTES, lhwndReadPipe As Long, lhwndWritePipe As Long
    Dim lBytesRead As Long, sBuffer As String
    Dim lPipeOutLen As Long, lPipeInLen As Long, lMaxInst As Long
    
    tSecurAttrib.nLength = Len(tSecurAttrib)
    tSecurAttrib.bInheritHandle = 1&
    tSecurAttrib.lpSecurityDescriptor = 0&

    lRetVal = CreatePipe(lhwndReadPipe, lhwndWritePipe, tSecurAttrib, 0)
    If lRetVal = 0 Then
        'CreatePipe failed
        Exit Function
    End If

    tStartupInf.cb = Len(tStartupInf)
    tStartupInf.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    tStartupInf.hStdOutput = lhwndWritePipe
    If bShowWindow Then
        'Show the DOS window
        tStartupInf.wShowWindow = SW_NORMAL
    Else
        'Hide the DOS window
        tStartupInf.wShowWindow = SW_HIDE
    End If

    lRetVal = CreateProcessA(0&, sCommandLine, tSecurAttrib, tSecurAttrib, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, tStartupInf, tProcInfo)
    If lRetVal <> 1 Then
        'CreateProcess failed
        Exit Function
    End If
    
    'Process created, wait for completion. Note, this will cause your application
    'to hang indefinately until this process completes.
    WaitForSingleObject tProcInfo.hProcess, INFINITE
    
    'Determine pipes contents
    lSuccess = GetNamedPipeInfo(lhwndReadPipe, PIPE_TYPE_BYTE, lPipeOutLen, lPipeInLen, lMaxInst)
    If lSuccess Then
        'Got pipe info, create buffer
        sBuffer = String(lPipeOutLen, 0)
        'Read Output Pipe
        lSuccess = ReadFile(lhwndReadPipe, sBuffer, lPipeOutLen, lBytesRead, 0&)
        If lSuccess = 1 Then
            'Pipe read successfully
            ShellExecuteCapture = Left$(sBuffer, lBytesRead)
            Form1.txtOutputs.Text = Form1.txtOutputs.Text & vbNewLine & Left$(sBuffer, lBytesRead) & vbNewLine
            Form1.txtOutputs.SelStart = Len(Form1.txtOutputs.Text)
        End If
    End If
    
    'Close handles
    Call CloseHandle(tProcInfo.hProcess)
    Call CloseHandle(tProcInfo.hThread)
    Call CloseHandle(lhwndReadPipe)
    Call CloseHandle(lhwndWritePipe)
End Function

Public Function ShellRun(sCmd As String, Timeout As Long) As String

Dim dStart As String
Dim sLine As String
Dim objShell As Object
Dim objExec As Object
Dim oOutput As Object

dStart = Now()
Set objShell = CreateObject("WScript.Shell")
Set objExec = objShell.Exec(sCmd)
Set oOutput = objExec.StdOut
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then ShellRun = ShellRun & sLine & vbCrLf
        If (Now() >= DateAdd("s", Timeout, dStart)) Then
           objExec.Terminate
           ShellRun = "serror"
           Exit Function
        End If
    Wend
Form1.txtOutputs.Text = Form1.txtOutputs.Text & vbNewLine & ShellRun & vbNewLine
Form1.txtOutputs.SelStart = Len(Form1.txtOutputs.Text)

End Function
