Attribute VB_Name = "mProcWait"
Option Explicit

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const WAIT_INFINITE = -1&

Private Type STARTUPINFO
  cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Long
  cbReserved2 As Long
  lpReserved2 As Long
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

Private Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessId As Long
  dwThreadID As Long
End Type

Private Declare Function CreateProcess Lib "kernel32" _
   Alias "CreateProcessA" _
  (ByVal lpAppName As Long, _
   ByVal lpCommandLine As String, _
   ByVal lpProcessAttributes As Long, _
   ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, _
   ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, _
   ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, _
   lpProcessInformation As PROCESS_INFORMATION) As Long
   
Private Declare Function WaitForSingleObject Lib "kernel32" _
  (ByVal hHandle As Long, _
   ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
  (ByVal hObject As Long) As Long
    
Public Sub RunProcess(cmdLine As String, OPT As Long)
   
   Select Case OPT
   Case 1
    RunProcessX cmdLine
   Case 2
    RunProcessE cmdLine
   End Select

End Sub

Private Sub RunProcessX(cmdLine As String)

   Dim proc As PROCESS_INFORMATION
   Dim start As STARTUPINFO
   
   start.cb = Len(start)
   
   Call CreateProcess(0&, cmdLine, 0&, 0&, 1&, _
                      NORMAL_PRIORITY_CLASS, 0&, 0&, _
                      start, proc)
   
End Sub

Private Sub RunProcessE(cmdLine As String)

   Dim proc As PROCESS_INFORMATION
   Dim start As STARTUPINFO
   
   start.cb = Len(start)
   
   Call CreateProcess(0&, cmdLine, 0&, 0&, 1&, _
                      NORMAL_PRIORITY_CLASS, 0&, 0&, _
                      start, proc)
   
   Call WaitForSingleObject(proc.hProcess, WAIT_INFINITE)
   
   Call CloseHandle(proc.hProcess)

   Call CloseHandle(proc.hThread)

End Sub



