Attribute VB_Name = "mFileFrmResource"
'Usage: FFResouce "FILEPATH" & "File.Name", FILE, "CUSTOM"

Public Enum AppResource
 RFA = 101
 RFB = 102
End Enum

Public Function FFResouce(resFILE As String, resID As AppResource, resTITLE As String) As String
    
    On Error GoTo ErrResouce
    Dim resBYTE() As Byte
    
    resBYTE = LoadResData(resID, resTITLE)
    
    Open resFILE For Binary Access Write As #1
    Put #1, , resBYTE
    Close #1
    FFResouce = resFILE
    Exit Function
    
ErrResouce:
    FFResouce = vbNullString
    MsgBox Err & ":Error in FFResouce.  Error Message: " & Err.Description, vbCritical, "Warning"
    
    Exit Function
    
End Function




