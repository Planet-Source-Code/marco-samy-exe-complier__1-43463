Attribute VB_Name = "mdIcons"
Option Explicit
Public Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public DefIcon As String
Function CountIcons(ByVal sFile As String)
Dim a%
On Error Resume Next
Dim lIcon
Do
lIcon = ExtractIcon(App.hInstance, sFile, a)
If lIcon = 0 Then Exit Do
a = a + 1
DestroyIcon lIcon
Loop
CountIcons = a
End Function
Function IconExtract(ByVal sFile As String, ByVal sIndex As Integer, ByVal DrawTo As Object, Optional ByVal x As Single = 0, Optional ByVal y As Single = 0)
Dim lIcon
lIcon = ExtractIcon(App.hInstance, sFile, sIndex)
DrawIcon DrawTo.hdc, x, y, lIcon
DestroyIcon lIcon
DrawTo.Refresh
End Function
Function AsExtract(ByVal sFile As String, ByVal DrawTo As Object, Optional ByVal x As Single = 0, Optional ByVal y As Single = 0)
Dim lIcon
lIcon = ExtractAssociatedIcon(App.hInstance, sFile, 0)
DrawIcon DrawTo.hdc, x, y, lIcon
DestroyIcon lIcon
DrawTo.Refresh
End Function
Function HasIcon(ByVal sFile As String) As Boolean
Dim lIcon
lIcon = ExtractIcon(App.hInstance, sFile, 0)
If Not lIcon = 0 Then
DestroyIcon lIcon
HasIcon = True
Else
HasIcon = False
End If
End Function
Function NormalizePath(sPath As String) As String
    If Right$(sPath, 1) <> "\" Then
        NormalizePath = sPath & "\"
    Else
        NormalizePath = sPath
    End If
End Function
Public Function GetIcon(sForm As Form, Optional ByVal sDefault As String) As String
'On Error Resume Next
DefIcon = sDefault
IconXT.Show vbModal, sForm
GetIcon = DefIcon
End Function
