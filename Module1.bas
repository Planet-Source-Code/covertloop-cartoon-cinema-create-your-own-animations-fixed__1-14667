Attribute VB_Name = "Module1"
'This module gets the short path name
'of a long path name

'Ex.  C:\Program Files\Script1.ccs
'       becomes
'     C:\Progra~1\Script1.ccs
Declare Function GetShortPathName Lib "kernel32" Alias _
  "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal _
  lpszShortPath As String, ByVal cchBuffer As Long) As Long


