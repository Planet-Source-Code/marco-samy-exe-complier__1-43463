VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   36
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Code Of the SFX "Win32.sfx" Which Located is the upper directory
'If the "Win32.sfx" Not Found in the upper directory, Please Compiles this
'project and rename it to "Win32.sfx" and but it in the Editor Directory
'////////////////////////////////////////////////////////////////////////
'/////////////////Created & Programmed By Marco Samy
'/////////////////marco_s2@hotmail.com
'////////////////////////////////////////////////////////////////////////
'/////////////////2002-2003   Egypt [El Minia]
'////////////////////////////////////////////////////////////////////////
Const BootStr = "Made By---- Marco Samy ----2002"
Dim MyFore As Single
Dim Msgs As New Collection
Private Sub Form_Load()
Collect1
Dim LastStr As String
Dim TmpCol As New Collection
GetAllAB Msgs.Item(Msgs.Count), ",", ",", TmpCol
BackColor = CSng(TmpCol(TmpCol.Count - 1))
MyFore = CSng(TmpCol(TmpCol.Count))
Show
WindowOnTop Me
DoEvents
StartAnim
ShowClose
Visible = False
DoEvents
Execute
End
End Sub
Function Collect1()
On Error GoTo Err1
Dim nf, ld As Byte, AllData As String, CPos As Single
nf = FreeFile
Open App.Path & "\" & App.EXEName & ".exe" For Binary As #nf
CPos = FileLen(App.Path & "\" & App.EXEName & ".exe")
While Not Left$(AllData, Len(BootStr)) = BootStr
Get #nf, CPos, ld
AllData = Chr(255 - ld) & AllData ' Simple Encryp
CPos = CPos - 1
Wend
Close #nf
GetAllAB AllData, vbCrLf, vbCrLf, Msgs
Exit Function
Err1:
Msgs.Add "Info File Not Found!"
Msgs.Add "Info File Not Found!"
End Function
Function StartAnim()
Dim CurMsg As String, StartC As Single, EndC As Single, cColor As Single, oTime As Long, nTime As Long
For I = 1 To Msgs.Count - 1
Cls
DoEvents
CurMsg = Msgs.Item(I)
StartC = MyFore
EndC = BackColor
oTime = timeGetTime
 For Z = 0 To 100 Step 5
 cColor = ColorPercentage(StartC, EndC, CSng(Z))
 ForeColor = cColor
 CurrentX = (Width - TextWidth(CurMsg)) / 2
 CurrentY = (Height - TextHeight(CurMsg)) / 2
 nTime = timeGetTime
 'If nTime - oTime < Val(100) Then '100* 20* 2 = 4000 (4 Seconds)
Count1:
 For G = 1 To 80 - (nTime - oTime): DoEvents: Next G: If timeGetTime - oTime < Val(80) Then GoTo Count1
 oTime = timeGetTime
 Print CurMsg
 'DoEvents
 Next Z
 For Z = 100 To 0 Step -5
 cColor = ColorPercentage(StartC, EndC, CSng(Z))
 ForeColor = cColor
 CurrentX = (Width - TextWidth(CurMsg)) / 2
 CurrentY = (Height - TextHeight(CurMsg)) / 2
Count2:
 For G = 1 To 80 - (nTime - oTime): DoEvents: Next G: If timeGetTime - oTime < Val(80) Then GoTo Count2
 oTime = timeGetTime
 Print CurMsg
 DoEvents
 Next Z
 For Z = 1 To 100
 DoEvents
 Next Z
Next I
End Function
Function ShowClose()
Dim CurMsg As String
CurMsg = "By Marco Samy"
AutoRedraw = False
Dim Hskip As Single, oHeight, oWidth
oWidth = Width: oHeight = Height
WindowState = vbNormal
Width = oWidth: Height = oHeight
Hskip = Height / Width
For I = Width To 800 Step -80
Move (Screen.Width - (Width - (Width - I))) / 2, (Screen.Height - (Height - (Hskip * (Width - I)))) / 2, Width - (Width - I), Height - (Hskip * (Width - I))
 cColor = ColorPercentage(vbBlack, vbBlue, CSng((I - 800) * 100 / (oWidth - 800)))
 ForeColor = cColor
 CurrentX = (Width - TextWidth(CurMsg)) / 2
 CurrentY = (Height - TextHeight(CurMsg)) / 2
 Print CurMsg
DoEvents
Next I
End Function
Function Execute()
On Error GoTo Err1
Dim Operate As String, FileName As String, Commands As String, Directory As String
Dim TmpCl As New Collection
GetAllAB Msgs.Item(Msgs.Count), ",", ",", TmpCl
Operate = Trim(TmpCl.Item(1))
If Operate = "Message" Then Exit Function
FileName = Trim(TmpCl.Item(2))
Commands = Trim(TmpCl.Item(3))
Directory = Trim(TmpCl.Item(3))
If Operate = "" Then Operate = "open"
If Directory = "" Then Directory = App.Path & "\"
If Not InStr(1, FileName, ":", vbTextCompare) = 0 Then FileName = FileName Else FileName = App.Path & "\" & FileName
ShellExecute Me.hwnd, Operate, FileName, Commands, Directory, vbNormalFocus
Err1:
End Function
