VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form IconXT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Selector"
   ClientHeight    =   5385
   ClientLeft      =   5955
   ClientTop       =   5130
   ClientWidth     =   6675
   Icon            =   "ExtractIcon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   60
      Width           =   5535
   End
   Begin MSComctlLib.ImageList ImTemp 
      Left            =   2640
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExtractIcon.frx":0ECA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ics 
      Left            =   3240
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.FileListBox FV 
      Height          =   675
      Left            =   2760
      TabIndex        =   11
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Help"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Tag             =   "#187,188"
      Top             =   4965
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Tag             =   "#40"
      Top             =   4965
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Tag             =   "#39"
      Top             =   4965
      Width           =   1455
   End
   Begin VB.PictureBox PicTemp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   545
      Left            =   3240
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   545
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   1830
      Left            =   2640
      TabIndex        =   3
      Top             =   590
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3228
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImTemp"
      ForeColor       =   -2147483640
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00C0E0FF&
      Height          =   1440
      Left            =   60
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.DriveListBox DV 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin MSComctlLib.ListView LV2 
      Height          =   2295
      Left            =   60
      TabIndex        =   0
      Top             =   2640
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   4048
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16744703
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Image Picture1 
      Height          =   735
      Left            =   5640
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Path :"
      Height          =   255
      Left            =   60
      TabIndex        =   14
      Tag             =   "#191"
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   5640
      TabIndex        =   13
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Icon Preview:"
      Height          =   255
      Left            =   5640
      TabIndex        =   12
      Tag             =   "#195"
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Files Contains Icons :"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Tag             =   "#193"
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Found Icons :"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Tag             =   "#194"
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Select File :"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Tag             =   "#192"
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "IconXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Dont As Boolean
Private Sub Command1_Click()
DefIcon = Text1.Text
Unload Me
End Sub





Private Sub Drive1_Change()

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
If Dont = True Then Exit Sub
PicTemp.BackColor = LV1.BackColor
ImTemp.MaskColor = PicTemp.BackColor
Dim Files As New Collection
Dim I
FV.Path = Dir1.Path
Dim x
x = 0
LV1.ListItems.Clear
Set LV1.Icons = Nothing

ImTemp.ListImages.Clear

For I = 0 To FV.ListCount - 1
x = x + 1
If HasIcon(NormalizePath(Dir1.Path) & FV.List(I)) = True Then
PicTemp.Cls
IconExtract NormalizePath(Dir1.Path) & FV.List(I), 0, PicTemp, 0, 0
ImTemp.ListImages.Add , , PicTemp.Image
Files.Add FV.List(I)
End If
Next I
If Not Files.Count = 0 Then Set LV1.Icons = ImTemp: LV1.View = lvwIcon Else LV1.View = lvwList: LV1.ListItems.Add , , "No Icons Found."
For I = 1 To Files.Count
LV1.ListItems.Add , , Files.Item(I), I
Next I
LV1_Click
End Sub

Private Sub DV_Change()
On Error GoTo devErr
Dir1.Path = DV.Drive
Exit Sub
devErr:
MsgBox "Error Reading from the device or the divce is not ready.", vbCritical
DV.Drive = Left$(Dir1.Path, 3)
End Sub

Private Sub Form_Load()
Dim I
If Dont = True Then Dont = False: Exit Sub
If Trim(DefIcon) = "" Then
Dir1_Change
Else
Dim FileName As String, FilePath As String, IconIndex As Integer
FileName = Between(DefIcon, "\", ",", 1)
FilePath = GetBL("\", DefIcon)
IconIndex = Val(GetAL(",", DefIcon))
Text1.Text = DefIcon
DefIcon = ""
'Show
DoEvents
Dont = True
If Len(FilePath) = 2 Then DV.Drive = FilePath & "\" Else DV.Drive = GetBF("\", FilePath, 1) & "\"
Dont = True
Dir1.Path = FilePath & "\"
Dont = False
Dir1_Change
DoEvents
For I = 1 To LV1.ListItems.Count
If LV1.ListItems(I).Text = FileName Then LV1.ListItems(I).Selected = True
Next I
LV1_Click
On Error Resume Next
LV2.ListItems(IconIndex + 1).Selected = True
LV2_Click
End If
End Sub

Private Sub LV1_Click()
If Not LV1.View = lvwIcon Then Exit Sub
PicTemp.BackColor = LV2.BackColor
Ics.MaskColor = PicTemp.BackColor
Dim I
LV2.ListItems.Clear
Set LV2.Icons = Nothing
Ics.ListImages.Clear
Dim ImC
ImC = CountIcons(NormalizePath(Dir1.Path) & LV1.SelectedItem.Text)
Label2.Caption = "Found Icons: " & ImC
For I = 1 To ImC
PicTemp.Cls
IconExtract NormalizePath(Dir1.Path) & LV1.SelectedItem.Text, I - 1, PicTemp, 0, 0
Ics.ListImages.Add , , PicTemp.Image
Next I
Set LV2.Icons = Ics
For I = 1 To ImC
LV2.ListItems.Add , , "Icon " & I, I
Next I
LV2_Click
End Sub

Private Sub LV1_KeyUp(KeyCode As Integer, Shift As Integer)
LV1_Click
End Sub

Private Sub LV2_Click()
Dim PicX As Picture
Set PicX = Ics.ListImages(LV2.SelectedItem.Index).ExtractIcon
Picture1.Picture = PicX
Picture1.Left = LV2.Left + LV2.Width + ((Width - (LV2.Left + LV2.Width) - Picture1.Width) / 2)
lblInfo.Caption = Picture1.Width / Screen.TwipsPerPixelX & " x " & Picture1.Height / Screen.TwipsPerPixelY
Text1.Text = NormalizePath(Dir1.Path) & LV1.SelectedItem.Text & "," & LV2.SelectedItem.Index - 1
Command1.Enabled = True
End Sub
Function ExtractRC(ByVal sIcon As String) As Picture
Dont = True
Load IconXT
Visible = False
Dim FileName As String, IconIndex As Integer
FileName = Trim(GetBF(",", sIcon, 1))
IconIndex = Val(Trim(GetAL(",", sIcon)))
PicTemp.BackColor = LV1.BackColor
ImTemp.MaskColor = PicTemp.BackColor
Set LV1.Icons = Nothing
ImTemp.ListImages.Clear
IconExtract FileName, IconIndex, PicTemp, 0, 0
ImTemp.ListImages.Add , , PicTemp.Image
Set ExtractRC = ImTemp.ListImages(1).ExtractIcon
ImTemp.ListImages.Clear

Unload IconXT
End Function
Function ExtractAS(ByVal sIcon As String) As Picture
Dont = True
Load IconXT
Visible = False
PicTemp.BackColor = LV1.BackColor
ImTemp.MaskColor = PicTemp.BackColor
Set LV1.Icons = Nothing
ImTemp.ListImages.Clear
AsExtract sIcon, PicTemp
ImTemp.ListImages.Add , , PicTemp.Image
Set ExtractAS = ImTemp.ListImages(1).ExtractIcon
ImTemp.ListImages.Clear
Unload IconXT
End Function
