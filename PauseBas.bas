Attribute VB_Name = "PauseBas"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Function Pause(interval)
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Function

