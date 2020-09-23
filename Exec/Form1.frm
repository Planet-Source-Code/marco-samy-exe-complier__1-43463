VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Executer Slide Show Editor"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8595
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList Export 
      Left            =   6480
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16711935
      _Version        =   393216
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Browse"
      Height          =   375
      Left            =   2400
      TabIndex        =   28
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Default"
      Height          =   375
      Left            =   960
      TabIndex        =   27
      Top             =   5880
      Width           =   1335
   End
   Begin VB.PictureBox PicTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   25
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox Fore 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7440
      MousePointer    =   10  'Up Arrow
      ScaleHeight     =   225
      ScaleWidth      =   1065
      TabIndex        =   24
      Top             =   4920
      Width           =   1095
   End
   Begin VB.PictureBox Back 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      MousePointer    =   10  'Up Arrow
      ScaleHeight     =   225
      ScaleWidth      =   1065
      TabIndex        =   22
      Top             =   4920
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   4800
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4920
      MaxLength       =   32
      TabIndex        =   20
      Text            =   "My Start"
      Top             =   4080
      Width           =   3615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Help"
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Close"
      Height          =   375
      Left            =   6600
      TabIndex        =   16
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Browse"
      Height          =   375
      Left            =   7080
      TabIndex        =   15
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Top             =   4560
      Width           =   7335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":030A
      Left            =   1200
      List            =   "Form1.frx":0317
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Messages :"
      Height          =   3375
      Left            =   100
      TabIndex        =   4
      Top             =   60
      Width           =   8415
      Begin VB.CommandButton Command6 
         Caption         =   "Edit"
         Height          =   375
         Left            =   7080
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Down"
         Height          =   375
         Left            =   4920
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Up"
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Height          =   2580
         IntegralHeight  =   0   'False
         Left            =   80
         TabIndex        =   5
         Top             =   720
         Width           =   8250
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "The File is in the same directory"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      ToolTipText     =   "Drag And Drop File Over Here (Any File *.*)"
      Top             =   3480
      Width           =   5775
   End
   Begin VB.Label Label9 
      Caption         =   "Choose 16 color Icon for smaller EXE size (The Icon Will Be Automatically Converted.)"
      Height          =   375
      Left            =   3840
      TabIndex        =   30
      Top             =   5880
      Width           =   3255
   End
   Begin VB.Label lblIco 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Icon"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   29
      Top             =   5520
      Width           =   7575
   End
   Begin VB.Label Label8 
      Caption         =   "Exe Icon :"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   5280
      Width           =   5415
   End
   Begin VB.Label Label7 
      Caption         =   "Font Color"
      Height          =   255
      Left            =   6000
      TabIndex        =   23
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Background"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Prog Title"
      Height          =   255
      Left            =   4080
      TabIndex        =   19
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "By Marco Samy -2002"
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   6480
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Commands"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Operation"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "File Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////////////
'/////////////////Created & Programmed By Marco Samy
'/////////////////marco_s2@hotmail.com
'////////////////////////////////////////////////////////////////////////
'/////////////////2002-2003   Egypt [El Minia]
'////////////////////////////////////////////////////////////////////////Const BootStr = "Made By---- Marco Samy ----2002"
Private Sub Back_Click()
Back.BackColor = GetColor(Me, Back.BackColor)
End Sub
Private Sub Command1_Click()
If Dir$(App.Path & "\Win32.sfx") = "" Then MsgBox "The Win32.sfx Module not found, See Compile SFX SRC\Project1.vbp For Details, Into (WIN32.SFX)", vbCritical: End
If Trim$(Text1.Text) = "" Then MsgBox "Select File First.", vbCritical: Exit Sub
On Error GoTo Err1:
Dim MyString As String, MyFile As String, MyData As String, MyCommand As String, Target As String
If Check1.Value = 1 Then MyFile = GetAL("\", Text1.Text) Else MyFile = Text1.Text
If Not Trim(Text2.Text) = "" Then MyCommand = Trim(Text2.Text) Else MyCommand = " "
MyData = Combo1.Text & "," & MyFile & "," & MyCommand & ", " & CStr(Back.BackColor) & "," & CStr(Fore.BackColor)
MyString = BootStr
For I = 0 To List1.ListCount - 1
MyString = MyString & vbCrLf & List1.List(I)
Next I
Target = Text3.Text
MyString = MyString & vbCrLf & MyData
FileCopy App.Path & "\Win32.sfx", GetBL("\", Text1.Text) & "\" & Target & ".exe"
Dim nf, Ab As Byte, oLen As Single
nf = FreeFile
oLen = FileLen(GetBL("\", Text1.Text) & "\" & Target & ".exe")
Open GetBL("\", Text1.Text) & "\" & Target & ".exe" For Binary As #nf
For I = 1 To Len(MyString)
Ab = 255 - Asc(Mid$(MyString, I, 1)) 'simple encryp
Put #nf, oLen + I, Ab
Next I
Close #nf
If lblIco.Caption = App.Path & "\Icosrc.srp,0" Then GoTo SkipIco
'Changing Exe Icon
oLen = 0
Dim Likes As Integer
REPEAT:
Dim StPoint As Long, AllByte As String, IcoByte As String, oIcoByte As String
AllByte = Space(FileLen(GetBL("\", Text1.Text) & "\" & Target & ".exe"))
IcoByte = Space(FileLen(App.Path & "\Icosrc.srp"))
Open GetBL("\", Text1.Text) & "\" & Target & ".exe" For Binary As #nf
Get #nf, , AllByte
Close #nf
Open App.Path & "\Icosrc.srp" For Binary As #nf
Get #nf, , IcoByte
Close #nf
oLen = 1
Export.ListImages.Clear
Export.ListImages.Add , , PicTemp.Image
SavePicture Export.ListImages.Item(1).ExtractIcon, App.Path & "\1.ico"
Export.ListImages.Clear
oLen = FileLen(App.Path & "\1.ico")
oIcoByte = Space(oLen)
Open App.Path & "\1.ico" For Binary As #nf
Get #nf, , oIcoByte
Close #nf
Kill App.Path & "\1.ico"
Likes = 0
For I = 1 To Len(IcoByte)
If Mid$(oIcoByte, I, 1) = Mid$(IcoByte, I, 1) Then Likes = Likes + 1 Else GoTo EXitCount
Next I
EXitCount:
IcoByte = Right$(IcoByte, Len(IcoByte) - Likes)
oIcoByte = Right$(oIcoByte, Len(oIcoByte) - Likes)
StPoint = InStr(1, AllByte, IcoByte, vbBinaryCompare)
If Not StPoint = 0 Then
Open GetBL("\", Text1.Text) & "\" & Target & ".exe" For Binary As #nf
Put #nf, StPoint, oIcoByte
Close #nf
GoTo REPEAT
End If
'/Changeing Icon
SkipIco:
If MsgBox("File Exporting Done, Do you want to launch it?", vbInformation + vbYesNo) = vbYes Then Shell GetBL("\", Text1.Text) & "\" & Target & ".exe", vbNormalFocus
Exit Sub
Err1:
MsgBox Err.Description, vbCritical
Target = "My Start"
Resume
End Sub
Private Sub Command10_Click()
PicTemp.Cls
AsExtract App.Path & "\Icosrc.srp", PicTemp
lblIco.Caption = App.Path & "\Icosrc.srp,0"
End Sub
Private Sub Command11_Click()
Dim new_Ic As String
new_Ic = GetIcon(Me, lblIco.Caption)
If Not new_Ic = "" Then
lblIco.Caption = new_Ic
PicTemp.Cls
IconExtract GetBL(",", new_Ic), Val(GetAL(",", new_Ic)), PicTemp
End If
End Sub
Private Sub Command2_Click()
Dim Idx
Idx = List1.ListCount
Dim Iq As String
Iq = InputBox("Add New Message" & vbCrLf & vbCrLf & "Write Your Message Text Here:", "New Message")
If Not Len(Iq) = 0 Then List1.AddItem Iq: List1.ListIndex = Idx
End Sub
Private Sub Command3_Click()
On Error Resume Next
List1.RemoveItem List1.ListIndex
End Sub
Private Sub Command4_Click()
On Error Resume Next
Dim Idx, Item
Idx = List1.ListIndex
If Idx = 0 Then Exit Sub
Item = List1.List(Idx)
List1.RemoveItem Idx
List1.AddItem Item, Idx - 1
List1.ListIndex = Idx - 1
End Sub
Private Sub Command5_Click()
Dim Idx, Item
Idx = List1.ListIndex
If Idx = List1.ListCount - 1 Then Exit Sub
Item = List1.List(Idx)
List1.RemoveItem Idx
List1.AddItem Item, Idx + 1
List1.ListIndex = Idx + 1
End Sub
Private Sub Command6_Click()
If List1.ListIndex < 0 Then Exit Sub
Dim Iq As String
Iq = InputBox("Add New Message" & vbCrLf & vbCrLf & "Write Your Message Text Here:", "New Message", List1.Text)
If Not Len(Iq) = 0 Then
Dim Idx
Idx = List1.ListIndex
List1.RemoveItem Idx
List1.AddItem Iq, Idx
List1.ListIndex = Idx
End If
End Sub
Private Sub Command7_Click()
On Error GoTo Err1
With Dlg
.CancelError = True
.FileName = Text1.Text
.ShowOpen
Text1.Text = .FileName
If HasIcon(.FileName) = True Then
lblIco.Caption = .FileName & ",0"
PicTemp.Cls
IconExtract .FileName, 0, PicTemp
End If
Dim SmPart As String
SmPart = GetBL("\", .FileName)
SmPart = GetAL("\", SmPart)
If Not SmPart = "" Then Text3.Text = SmPart
End With
Err1:
End Sub
Private Sub Command8_Click()
Unload Me
End Sub
Private Sub Command9_Click()
Dim Msg As String
Msg = "Create You Slide Show Messages By Useing Add, Remove, Edit Buttons" & vbCrLf & "Select Your File(Any File Type Associated with any program or Any Executable.)" & vbCrLf & "Click Create.  " & vbCrLf & vbCrLf & "Created & Programmed By : Marco Samy Nasif" & vbCrLf & "marco_s2@hotmail.com" & vbCrLf & "Call (+20) 0107242974"
MsgBox Msg, vbInformation
End Sub
Private Sub Fore_Click()
Fore.BackColor = GetColor(Me, Fore.BackColor)
End Sub
Private Sub Form_Load()
PicTemp.Width = 32 * Screen.TwipsPerPixelX: PicTemp.Height = 32 * Screen.TwipsPerPixelY
If App.PrevInstance = True Then MsgBox "Another Window of this program is already running.", vbCritical: End
Combo1.Text = "open"
AsExtract App.Path & "\Icosrc.srp", PicTemp
lblIco.Caption = App.Path & "\Icosrc.srp,0"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Iq
Iq = MsgBox("Are you sure you want to exit?", vbExclamation + vbYesNo)
If Iq = vbYes Then
End
Else
Cancel = 1
End If
End Sub
Private Sub PicTemp_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If Data.GetFormat(15) = True Then
For I = 1 To Data.Files.Count
If HasIcon(Data.Files(I)) = True Then
lblIco.Caption = Data.Files(I) & ",0"
PicTemp.Cls
IconExtract Data.Files(I), 0, PicTemp
GoTo EndI
End If
Next I
EndI:
End If
End Sub
Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If Data.GetFormat(15) = True Then
Text1.Text = Data.Files(1)
For I = 1 To Data.Files.Count
If HasIcon(Data.Files(I)) = True Then
lblIco.Caption = Data.Files(I) & ",0"
PicTemp.Cls
IconExtract Data.Files(I), 0, PicTemp
GoTo EndI
End If
Next I
EndI:
Dim SmPart As String
SmPart = GetBL("\", Data.Files(1))
SmPart = GetAL("\", Data.Files(1))
If Not SmPart = "" Then Text3.Text = SmPart
End If
End Sub
