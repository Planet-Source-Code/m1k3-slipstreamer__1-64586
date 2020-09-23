Attribute VB_Name = "mAPI"
Option Explicit

Global sCDRM As String
Global sCOPY As String
Global sSPEX As String
Global sSVCP As String
Global sFHND As String
Global sSPNM As String

Global FSIZE As Long
Global TSIZE As Long

Public Const oSetup = "\setup.exe /a"
Public Const SWA = " /a "
Public Const SWT = " /t:"
Public Const SWC = " /c"
Public Const SWX = " /x:"
Public Const SWS = " -s:"
Public Const i386 = "i386\Update\"

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
 (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
  ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub MakeDir(ByVal pthSPEC As String)

    Dim pos As Integer
    Dim strTemp As String
    pos = InStrRev(pthSPEC, Chr(92))

    If pos < InStrRev(pthSPEC, ".") Then
        pthSPEC = Left$(pthSPEC, pos)
    ElseIf Right$(pthSPEC, 1) <> Chr(92) Then pthSPEC = pthSPEC & Chr(92)
    End If

    If Left$(pthSPEC, 2) = "\\" Then
        pos = InStr(InStr(3, pthSPEC, Chr(92)) + 1, pthSPEC, Chr(92))
    Else: pos = 1
    End If
    
    Let pos = InStr(pos, pthSPEC, Chr(92))

    Do While pos <> 0
        pos = InStr(pos + 1, pthSPEC, Chr(92))

        If pos <> 0 Then
            strTemp = Left$(pthSPEC, pos)
            If Dir$(strTemp, vbDirectory) = Empty Then MkDir strTemp
        End If
    Loop
    
End Sub

