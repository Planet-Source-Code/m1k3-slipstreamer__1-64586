Attribute VB_Name = "mDrives"
Option Explicit

Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias _
 "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
  ByVal lpBuffer As String) As Long
  
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
 (ByVal nDrive As String) As Long

Public Function IsDriveReady(drvSPEC As String) As Boolean

   Dim m_oFSO As New FileSystemObject
   Set m_oFSO = New FileSystemObject
     
    IsDriveReady = m_oFSO.GetDrive(drvSPEC).IsReady

   Set m_oFSO = Nothing
   
End Function

Private Function Enumerate_Drives() As String

   Dim result As String
   
   result = String(255, Chr$(0))
   GetLogicalDriveStrings 255, result
   Enumerate_Drives = result
   
End Function

Public Function Get_CdrList(ByVal cb As ComboBox)
 
 Dim i As Long
 Dim arrDrives() As String
  arrDrives = Split(Enumerate_Drives, Chr$(0))
  
 For i = LBound(arrDrives) To UBound(arrDrives)
  If arrDrives(i) <> Empty And getType(arrDrives(i)) = 5 Then
   If IsDriveReady(arrDrives(i)) Then
    cb.Clear
    cb.AddItem arrDrives(i) & GetDriveLabel(arrDrives(i))
    cb.Text = cb.List(0)
   Else
    cb.Clear
    cb.AddItem arrDrives(i) & "      <NO CD>"
    cb.Text = cb.List(0)
   End If
  End If
 Next

End Function

Public Function Get_HddList(ByVal ls As ListBox)
 
 Dim i As Long
 Dim arrDrives() As String
  arrDrives = Split(Enumerate_Drives, Chr$(0))
  
 For i = LBound(arrDrives) To UBound(arrDrives)
  If arrDrives(i) <> Empty And getType(arrDrives(i)) = 3 Then
    ls.AddItem arrDrives(i) & vbTab & DriveFreeSpace(arrDrives(i))
  End If
 Next

End Function

Private Function getType(drvSPEC As String) As Long

    Select Case GetDriveType(drvSPEC)
        Case 2
           getType = 2 'RMV
        Case 3
           getType = 3 'HDD
        Case Is = 4
           getType = 4 'REM
        Case Is = 5
           getType = 5 'CDR
        Case Is = 6
           getType = 6 'RAM
        Case Else
           getType = 7 'UNK
    End Select
    
End Function

Public Function GetDriveLabel(drvSPEC As String) As String

 On Error Resume Next
 
 Dim m_oFSO As New FileSystemObject
  GetDriveLabel = m_oFSO.Drives(Left(drvSPEC, 1)).VolumeName
 Set m_oFSO = Nothing

End Function

Public Function DriveFreeSpace(drvPATH As String) As String

    Dim m_oFSO As New FileSystemObject
    Dim sfFolder As Scripting.Folder
    Dim sfDrive As Scripting.Drive
    Dim Gigabyte As Long
    Dim Megabyte As Long
    Dim lngGig As Double
    Dim lngMegs As Double
    Dim dblTemp As Double
    
    On Error GoTo Procedure_Err
    
    Megabyte = 1048576
    Gigabyte = 1073741824

    If Trim(drvPATH) <> Empty Then
        Set m_oFSO = New Scripting.FileSystemObject
        Set sfFolder = m_oFSO.GetFolder(drvPATH)
        Set sfDrive = sfFolder.Drive
        lngGig = (sfDrive.FreeSpace / Gigabyte)
        lngGig = Int(lngGig)
        dblTemp = lngGig * Gigabyte

        If dblTemp < sfDrive.FreeSpace Then
            lngMegs = sfDrive.FreeSpace - dblTemp
            lngMegs = (lngMegs / Megabyte)
        Else
            lngMegs = (sfDrive.FreeSpace / Megabyte)
        End If
        lngMegs = Int(lngMegs)
        
        If lngGig > 0 Then
         DriveFreeSpace = lngGig & Chr(32) & Chr(45) & Chr(32) & "Gigabyte(s) Free"
        Else
         DriveFreeSpace = lngMegs & Chr(32) & Chr(45) & Chr(32) & "Megabyte(s) Free"
        End If
        
    End If
    Set m_oFSO = Nothing
Procedure_Exit:
    Exit Function
Procedure_Err:
    MsgBox Err.Description
    Resume Procedure_Exit
    Resume
    
End Function
