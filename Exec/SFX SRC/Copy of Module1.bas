Attribute VB_Name = "API1"
'/////////////Sytem API Functions for Magic Copy////////////
'/////////////By Marco Samy Nasif 2002//////////////////////
'///////////////////////////////////////////////////////////
Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Function WindowOnTop(ByVal sWnd As Form) As Integer
Dim retVal
retVal = SetWindowPos(sWnd.hwnd, -1, sWnd.Left / Screen.TwipsPerPixelX, sWnd.Top / Screen.TwipsPerPixelY, sWnd.Width / Screen.TwipsPerPixelX, sWnd.Height / Screen.TwipsPerPixelY, &H20)
WindowOnTop = retVal
End Function
Function WindowNoTop(ByVal sWnd As Form) As Integer
Dim retVal
retVal = SetWindowPos(sWnd.hwnd, -2, sWnd.Left / Screen.TwipsPerPixelX, sWnd.Top / Screen.TwipsPerPixelY, sWnd.Width / Screen.TwipsPerPixelX, sWnd.Height / Screen.TwipsPerPixelY, &H20)
WindowNoTop = retVal
End Function
