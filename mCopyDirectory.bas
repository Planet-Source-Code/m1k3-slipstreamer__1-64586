Attribute VB_Name = "mCopyDirectory"
Option Explicit

Private Const MAX_PATH = 260

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
 (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
 
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
 (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
 
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Declare Function GetLastError Lib "kernel32" () As Long

Private Const ERROR_NO_MORE_FILES = 18&
Private Const INVALID_HANDLE_VALUE = -1
Private Const DDL_DIRECTORY = &H10

Private Function CopyFiles(ByVal F_DIR As String, ByVal T_DIR As String) As Long

Dim nFiles As Long
Dim cDirs As Collection
Dim sFile As String
Dim nSearch As Long
Dim wFile As WIN32_FIND_DATA
Dim i As Integer

    Set cDirs = New Collection
    nSearch = FindFirstFile( _
        F_DIR & "*.*", wFile)
    If nSearch <> INVALID_HANDLE_VALUE Then
        Do
            sFile = wFile.cFileName
            sFile = Left$(sFile, InStr(sFile, Chr$(0)) - 1)
            If sFile <> "." And sFile <> ".." Then
                nFiles = nFiles + 1
                If wFile.dwFileAttributes And DDL_DIRECTORY Then
                    MkDir T_DIR & sFile
                    cDirs.Add sFile
                Else
                    FileCopy F_DIR & sFile, T_DIR & sFile
                End If
            End If
            sFHND = Mid(F_DIR, 4, (Len(F_DIR) - 3)) & sFile
            DoEvents
            If FindNextFile(nSearch, wFile) = 0 Then Exit Do
        Loop
        FindClose nSearch
    End If

    For i = 1 To cDirs.Count
        sFile = cDirs(i)
        nFiles = nFiles + CopyFiles(F_DIR & sFile & Chr(92), T_DIR & sFile & Chr(92))
    Next i

    CopyFiles = nFiles
    
End Function

Public Function XCopyFile(F_DIR As String, T_DIR As String) As Long

Dim nFiles As Long

    If GetAttr(F_DIR) And vbDirectory Then
        If Right$(F_DIR, 1) <> Chr(92) Then F_DIR = F_DIR & Chr(92)
        If Right$(T_DIR, 1) <> Chr(92) Then T_DIR = T_DIR & Chr(92)

        On Error Resume Next
        MkDir T_DIR
        If Err.Number = 0 Then nFiles = 1
        On Error GoTo 0
        nFiles = nFiles + CopyFiles(F_DIR, T_DIR)
    Else
        FileCopy F_DIR, T_DIR
        nFiles = 1
    End If

    XCopyFile = nFiles
    
End Function

Public Function GetPathSize(ByRef pthSPEC As String) As Double

 Dim m_oFSO As FileSystemObject
 Dim target_folder As Folder

    DoEvents

    Set m_oFSO = New FileSystemObject
    Set target_folder = m_oFSO.GetFolder(pthSPEC)
    GetPathSize = target_folder.Size
    Set m_oFSO = Nothing
    
End Function



