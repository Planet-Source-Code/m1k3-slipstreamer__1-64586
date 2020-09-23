Attribute VB_Name = "mLinkLabel"
Public Enum OpType
    Startup = 1
    Click = 2
    FormMove = 3
    LinkMove = 4
End Enum

Dim Clicked As Boolean

'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
 '(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
 ' ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) _
 ' As Long

Public Sub MakeLink(LabelName As Label, Operation As OpType, Optional FormName As Form, Optional ExLink As String)
    
    Dim Openpage As Integer

    Select Case Operation
        Case LinkMove
        LabelName.ForeColor = 16711680
        LabelName.FontUnderline = True
        Case Click
        Openpage = ShellExecute(FormName.hwnd, "Open", ExLink, "", App.Path, 1)
        LabelName.ForeColor = 8388736
        Clicked = True
        Case FormMove
        LabelName.FontUnderline = False

         If Not Clicked Then
            LabelName.ForeColor = 16711680
         Else
            LabelName.ForeColor = 16711680 '8388736
         End If
        
        Case Startup
        LabelName.ForeColor = 16711680
    End Select
End Sub
