VERSION 5.00
Begin VB.Form ColorSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Selector"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   Icon            =   "Form1c.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1c.frx":030A
   ScaleHeight     =   334
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   396
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   17
      Left            =   4560
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   35
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   16
      Left            =   5040
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   34
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   15
      Left            =   5520
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   33
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   14
      Left            =   5520
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   31
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   13
      Left            =   5040
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   30
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   12
      Left            =   4560
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   29
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   5520
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   28
      Top             =   2160
      Width           =   375
   End
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   5040
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   27
      Top             =   2160
      Width           =   375
   End
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   4560
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   26
      Top             =   2160
      Width           =   375
   End
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   5520
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   25
      Top             =   1680
      Width           =   375
   End
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   5040
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   24
      Top             =   1680
      Width           =   375
   End
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   4560
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   23
      Top             =   1680
      Width           =   375
   End
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   5520
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   22
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   5040
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   21
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   4560
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   20
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   5520
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   19
      Top             =   720
      Width           =   375
   End
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   5040
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   18
      Top             =   720
      Width           =   375
   End
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   4560
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   17
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Screen"
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   4080
      Width           =   1335
   End
   Begin VB.PictureBox EF4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3000
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   12
      Top             =   3960
      Width           =   255
   End
   Begin VB.PictureBox EF3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2640
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   11
      Top             =   3960
      Width           =   255
   End
   Begin VB.PictureBox EF2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2280
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   10
      Top             =   3960
      Width           =   255
   End
   Begin VB.PictureBox EF1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   1920
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   9
      Top             =   3960
      Width           =   255
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   840
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "0"
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "0"
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "0"
      Top             =   3960
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3360
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   2
      Top             =   3960
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3870
      Left            =   3960
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   1
      Top             =   45
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   45
      MouseIcon       =   "Form1c.frx":0BD4
      Picture         =   "Form1c.frx":149E
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   45
      Width           =   3840
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   240
         Left            =   0
         Picture         =   "Form1c.frx":314E0
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   495
      Left            =   4560
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   300
      X2              =   258
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "By Marco Samy 10/2002"
      ForeColor       =   &H80000011&
      Height          =   495
      Left            =   4560
      TabIndex        =   32
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Prefers :"
      Height          =   255
      Left            =   4560
      TabIndex        =   15
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Color B"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Color G"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Color R"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
End
Attribute VB_Name = "ColorSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Oy As Single
Dim Dont As Boolean
Private Sub Command1_Click()
DefColor = Picture3.BackColor
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim dc1
dc1 = GetDC(0)
Load frmCap2
frmCap2.Move 0, 0, Screen.Width, Screen.Height
BitBlt frmCap2.hdc, 0, 0, Screen.Width, Screen.Height, dc1, 0, 0, vbSrcCopy
frmCap2.Refresh
WindowOnTop frmCap2
frmCap2.Show vbModal, Me
End Sub

Private Sub EF1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Hell
If Button = 1 Then
Picture3.BackColor = EF1.Point(x, y)
Dim R, G, B
ColorRGB Picture3.BackColor, R, G, B
Text1.Text = R
Text2.Text = G
Text3.Text = B
End If
Hell:
End Sub
Private Sub EF2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Hell
If Button = 1 Then
Picture3.BackColor = EF2.Point(x, y)
Dim R, G, B
ColorRGB Picture3.BackColor, R, G, B
Text1.Text = R
Text2.Text = G
Text3.Text = B
End If
Hell:
End Sub
Private Sub EF3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Hell
If Button = 1 Then
Picture3.BackColor = EF3.Point(x, y)
Dim R, G, B
ColorRGB Picture3.BackColor, R, G, B
Text1.Text = R
Text2.Text = G
Text3.Text = B

End If
Hell:
End Sub

Private Sub EF4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Hell
If Button = 1 Then
Picture3.BackColor = EF4.Point(x, y)
Dim R, G, B
ColorRGB Picture3.BackColor, R, G, B
Text1.Text = R
Text2.Text = G
Text3.Text = B
End If
Hell:
End Sub

Private Sub Form_Load()
Oy = 1

Dim R, G, B
ColorRGB DefColor, R, G, B
Text1.Text = R
Text2.Text = G
Text3.Text = B
Text1_KeyUp 0, 0
'Show
'For R = 0 To 255
'For G = 0 To 255
'Picture1.PSet (R, G), RGB(R, G, (255 / 2))
'Next G
'DoEvents
'Next R
End Sub

Private Sub PF_Click(Index As Integer)
Dim R, G, B
Shape1.Left = PF(Index).Left - ((Shape1.Width - PF(Index).Width) / 2) + 0.25
Shape1.Top = PF(Index).Top - ((Shape1.Height - PF(Index).Height) / 2) + 0.25
If Shape1.Visible = False Then Shape1.Visible = True
ColorRGB PF(Index).BackColor, R, G, B
Text1.Text = R
Text2.Text = G
Text3.Text = B
Text1_KeyUp 0, 0
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.Visible = False
Dim rc As RECT
GetWindowRect Picture1.hwnd, rc
ClipCursor rc
Picture2.AutoRedraw = False
Dim R, G, B
ColorRGB Picture1.Point(x, y), R, G, B
For B = 0 To 255
Picture2.Line (0, B)-(Picture2.Width, B), RGB(R, G, B), B
Next B
Picture2_MouseDown Button, Shift, 5, Oy
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Hell
If Button = 1 Then
Dim R, G, B
ColorRGB Picture1.Point(x, y), R, G, B
For B = 0 To 255
Picture2.Line (0, B)-(Picture2.Width, B), RGB(R, G, B), B
Next B
End If
Picture2_MouseMove Button, Shift, 5, Oy
Hell:
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim rc As RECT
GetWindowRect GetDesktopWindow, rc
ClipCursor rc

Picture2.AutoRedraw = True
Dim R, G, B
ColorRGB Picture1.Point(x, y), R, G, B
Image1.Move x - (Image1.Width / 2), y - (Image1.Width / 2)
Image1.Visible = True
For B = 0 To 255
Picture2.Line (0, B)-(Picture2.Width, B), RGB(R, G, B), B
Next B
Picture2_MouseUp Button, Shift, 5, Oy
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Hell
EF1.AutoRedraw = False
EF2.AutoRedraw = False
EF3.AutoRedraw = False
EF4.AutoRedraw = False
Picture3.AutoRedraw = False
Picture3.BackColor = Picture2.Point(x, y)
Dim R, G, B
ColorRGB Picture3.BackColor, R, G, B
Text1.Text = R
Text2.Text = G
Text3.Text = B
Dim R1, G1, B1
ColorRGB GrayColor(Picture3.BackColor), R1, G1, B1
For I = 0 To EF1.Height
EF1.Line (0, I)-(EF1.Width, I), ColorRGBPercentX(R, G, B, R1, G1, B1, I / EF1.Height * 100)
Next I
ColorRGB NegativeColor(Picture3.BackColor), R1, G1, B1
For I = 0 To EF1.Height
EF2.Line (0, I)-(EF1.Width, I), ColorRGBPercentX(R, G, B, R1, G1, B1, I / EF1.Height * 100)
Next I
R1 = 0: G1 = 0: B1 = 0
For I = 0 To EF1.Height
EF3.Line (0, I)-(EF1.Width, I), ColorRGBPercentX(R, G, B, R1, G1, B1, I / EF1.Height * 100)
Next I
R1 = 255: G1 = 255: B1 = 255
For I = 0 To EF1.Height
EF4.Line (0, I)-(EF1.Width, I), ColorRGBPercentX(R, G, B, R1, G1, B1, I / EF1.Height * 100)
Next I
Hell:
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Hell
If Button = 1 Then
Picture3.BackColor = Picture2.Point(x, y)
Dim R, G, B
ColorRGB Picture3.BackColor, R, G, B
Dim R1, G1, B1
ColorRGB GrayColor(Picture3.BackColor), R1, G1, B1
For I = 0 To EF1.Height
EF1.Line (0, I)-(EF1.Width, I), ColorRGBPercentX(R, G, B, R1, G1, B1, I / EF1.Height * 100)
Next I
ColorRGB NegativeColor(Picture3.BackColor), R1, G1, B1
For I = 0 To EF1.Height
EF2.Line (0, I)-(EF1.Width, I), ColorRGBPercentX(R, G, B, R1, G1, B1, I / EF1.Height * 100)
Next I
R1 = 0: G1 = 0: B1 = 0
For I = 0 To EF1.Height
EF3.Line (0, I)-(EF1.Width, I), ColorRGBPercentX(R, G, B, R1, G1, B1, I / EF1.Height * 100)
Next I
R1 = 255: G1 = 255: B1 = 255
For I = 0 To EF1.Height
EF4.Line (0, I)-(EF1.Width, I), ColorRGBPercentX(R, G, B, R1, G1, B1, I / EF1.Height * 100)
Next I
End If
Hell:
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Hell
EF1.AutoRedraw = True
EF2.AutoRedraw = True
EF3.AutoRedraw = True
EF4.AutoRedraw = True
Picture3.AutoRedraw = True
Picture3.BackColor = Picture2.Point(x, y)
Dim R, G, B
ColorRGB Picture3.BackColor, R, G, B
If Not Dont = True Then
Text1.Text = R
Text2.Text = G
Text3.Text = B
End If
Dim R1, G1, B1
ColorRGB GrayColor(Picture3.BackColor), R1, G1, B1
For I = 0 To EF1.Height
EF1.Line (0, I)-(EF1.Width, I), ColorRGBPercentX(R, G, B, R1, G1, B1, I / EF1.Height * 100)
Next I
ColorRGB NegativeColor(Picture3.BackColor), R1, G1, B1
For I = 0 To EF1.Height
EF2.Line (0, I)-(EF1.Width, I), ColorRGBPercentX(R, G, B, R1, G1, B1, I / EF1.Height * 100)
Next I
R1 = 0: G1 = 0: B1 = 0
For I = 0 To EF1.Height
EF3.Line (0, I)-(EF1.Width, I), ColorRGBPercentX(R, G, B, R1, G1, B1, I / EF1.Height * 100)
0 Next I
R1 = 255: G1 = 255: B1 = 255
For I = 0 To EF1.Height
EF4.Line (0, I)-(EF1.Width, I), ColorRGBPercentX(R, G, B, R1, G1, B1, I / EF1.Height * 100)
Next I
Oy = y
Line1.Y1 = Picture2.Top + Oy
Line1.Y2 = Picture2.Top + Oy
Hell:
End Sub

Public Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If Dont = True Then Exit Sub
Dont = True
Text1.Text = Val(Text1.Text)
If Val(Text1.Text) > 255 Then Text1.Text = 255
Picture1_MouseUp 1, 0, Val(Text1.Text), Val(Text2.Text) ' - 1
Picture2_MouseUp 1, 0, 5, Val(Text3.Text) ' - 1
Dont = False
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
If Dont = True Then Exit Sub
Dont = True
Text2.Text = Val(Text2.Text)
If Val(Text2.Text) > 255 Then Text2.Text = 255
Picture1_MouseUp 1, 0, Val(Text1.Text), Val(Text2.Text) '- 1
Picture2_MouseUp 1, 0, 5, Val(Text3.Text) '- 1
Dont = False
End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
If Dont = True Then Exit Sub
Dont = True
Text3.Text = Val(Text3.Text)
If Val(Text3.Text) > 255 Then Text3.Text = 255
Picture1_MouseUp 1, 0, Val(Text1.Text), Val(Text2.Text) ' - 1
Picture2_MouseUp 1, 0, 5, Val(Text3.Text) '- 1
Dont = False
End Sub
