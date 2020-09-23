Attribute VB_Name = "Module2"
'API Functions Needed By Slide Show sfx
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
