VERSION 5.00
Begin VB.Form frmCap2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form2"
   MouseIcon       =   "frmCap2.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCap2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim R, G, B
ColorRGB Point(x, y), R, G, B
ColorSelect.Text1.Text = R
ColorSelect.Text2.Text = G
ColorSelect.Text3.Text = B
ColorSelect.Text1_KeyUp 0, 0
Unload Me
End Sub
