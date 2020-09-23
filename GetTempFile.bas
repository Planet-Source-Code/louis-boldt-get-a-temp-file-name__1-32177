Attribute VB_Name = "modGetTempFile"
Option Explicit
' Louis Boldt 2/22/2002
' Usage chose one depending on you needs
'
'  sTempFile = fnGetTempPath(sPrefix, sDir)
'  sTempFile = fnGetTempPath(sPrefix)
'  sTempFile = fnGetTempPath(, sDir)
'  sTempFile = fnGetTempPath()
'  sTempFile = fnGetTempPath

' Parms (both are optional)
'  sPrefix: Use if you want the file name to start
'  with your defined characters. Only first 3 chars are used
'  if ommited file is named @@@@.tmp
'  sDir: if you want a temp file in a specfic directory
'  if omitted the system temp directory is used.
'
' Returns
'  The name of the empty file created.

' get the directory in the temp enviromental variable
' dir name  is returned in lpBuffer
Public Declare Function GetTempDir Lib "kernel32" _
      Alias "GetTempPathA" _
      (ByVal nBufferLength As Long, _
       ByVal lpBuffer As String) As Long
       
' get the name and allocate the file
' path
Public Declare Function GetTempFilename Lib "kernel32" _
      Alias "GetTempFileNameA" _
      (ByVal lpszPath As String, _
       ByVal lpPrefixString As String, _
       ByVal wUnique As Long, _
       ByVal lpTempFileName As String) As Long
       
Private Const MAX_PATH As Long = 256
'
' if path is passed it is used else the temp path is used
' returns temp path &  temp file name
' Can be optionaly be passed a filename Prefix
' ----------------------------------------------------------------------
Public Function fnGetTempPath( _
                Optional sOptPrefix As String = "", _
                Optional sOptDir As String = "") As String
' ----------------------------------------------------------------------
  Dim sRoutineID As String

  Dim lReturnVal As Long
  ' be sure variables are long enought
  ' MAX_PATH  is used in API call also
  Dim sTempDir As String * MAX_PATH
  Dim sTempFilename As String * MAX_PATH
  
  On Error GoTo Handle_Error
  sRoutineID = "modGetTempFile  fnGetTempPath"
  
  ' be sure the string ends with a null char
  ' or error can result

  sOptPrefix = Trim$(sOptPrefix) & vbNullChar
 
  ' get the path in the temp enviromental variable
  ' or use the optional path ( check for ending \)
  sOptDir = Trim$(sOptDir)
  If Len(sOptDir) < 2 Then
    lReturnVal = GetTempDir(MAX_PATH, sTempDir)
  Else
    If Right$(sOptDir, 1) <> "\" Then
      sOptDir = sOptDir & "\"
    End If
    sTempDir = sOptDir & vbNullChar
  End If
  
  ' now get and create a temp file in that path
  lReturnVal = GetTempFilename(sTempDir, sOptPrefix, 0, sTempFilename)
  'return the name of the file
  fnGetTempPath = sTempFilename

Exit_Point:
  Exit Function

Handle_Error:

  MsgBox "Unexpected Error Returned" & vbCrLf _
       & "At   " & sRoutineID & vbCrLf _
       & "Nmbr " & Err.Number & vbCrLf _
       & "Desc " & Err.Description & vbCrLf _
       & "Srce " & Err.Source & vbCrLf _
       & "Time " & Now & vbCrLf _
       & "Path " & App.Path, _
     16, "You have a Problem!"

  fnGetTempPath = ""
  Resume Exit_Point
End Function
