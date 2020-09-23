Attribute VB_Name = "ModCOlors"
'Color Controller Module
'Copyright (c) Marco Samy 2002
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function ClipCursor Lib "user32.dll" (lpRect As RECT) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public DefColor As Single
Function ColorRGB(ByVal sColor As Single, ByRef ColorR, ByRef ColorG, ByRef ColorB)
    ColorR = (sColor And 255)
    ColorG = (sColor And 65280) / 256
    ColorB = (sColor And 16711680) / 65536
End Function
Function RndColor()
Dim R, G, B
Randomize
R = Int(Rnd * 256)
Randomize
G = Int(Rnd * 256)
Randomize
B = Int(Rnd * 256)
RndColor = RGB(R, G, B)
End Function
Function ColorPercentage(sColor1, sColor2, Optional sPercentTo1 As Single = 50)
Dim vR, vB, vG, vR1, vB1, vG1, vR2, vB2, vG2
ColorRGB sColor1, vR, vG, vB
ColorRGB sColor2, vR1, vG1, vB1
If Val(sPercentTo1) > 100 Then sPercentTo1 = 1 Else sPercentTo1 = sPercentTo1 / 100
vR2 = (vR * sPercentTo1) + (vR1 * (1 - sPercentTo1))
vG2 = (vG * sPercentTo1) + (vG1 * (1 - sPercentTo1))
vB2 = (vB * sPercentTo1) + (vB1 * (1 - sPercentTo1))
ColorPercentage = RGB(vR2, vG2, vB2)
End Function
Function BitPercent(sBit1, sBit2, Optional sPercentTo1 As Single = 50) As Byte
BitPercent = (sBit1 * sPercentTo1) + (sBit2 * (1 - sPercentTo1))
End Function
Function ColorRGBPercent(ByVal vR1, ByVal vB1, ByVal vG1, ByVal vR2, ByVal vB2, ByVal vG2, ByRef vR, ByRef vB, ByRef vG, Optional sPercentTo1 As Single = 50)
If Val(sPercentTo1) > 100 Then sPercentTo1 = 1 Else sPercentTo1 = sPercentTo1 / 100
vR = (vR1 * sPercentTo1) + (vR2 * (1 - sPercentTo1))
vG = (vG1 * sPercentTo1) + (vG2 * (1 - sPercentTo1))
vB = (vB1 * sPercentTo1) + (vB2 * (1 - sPercentTo1))
End Function
Function ColorRGBPercentX(ByVal vR1, ByVal vB1, ByVal vG1, ByVal vR2, ByVal vB2, ByVal vG2, Optional sPercentTo1 As Single = 50)
If Val(sPercentTo1) > 100 Then sPercentTo1 = 1 Else sPercentTo1 = sPercentTo1 / 100
ColorRGBPercentX = RGB((vR1 * sPercentTo1) + (vR2 * (1 - sPercentTo1)), (vG1 * sPercentTo1) + (vG2 * (1 - sPercentTo1)), (vB1 * sPercentTo1) + (vB2 * (1 - sPercentTo1)))
End Function
Function NegativeColor(sColor)
Dim vR, vB, vG
ColorRGB sColor, vR, vG, vB
vR = 255 - vR
vG = 255 - vG
vB = 255 - vB
NegativeColor = RGB(vR, vG, vB)
End Function
Function GrayColor(sColor)
Dim vR, vB, vG, vMid
ColorRGB sColor, vR, vG, vB
vMid = (0.5 + (0.299 * vR) + (0.587 * vG) + (0.114 * vB))
vR = IIf(Value >= 256, 255, vMid)
vG = IIf(Value >= 256, 255, vMid)
vB = IIf(Value >= 256, 255, vMid)
GrayColor = RGB(vR, vG, vB)
End Function
Function GetColor(sForm As Form, Optional ByVal sDef As Single) As Single
DefColor = sDef
ColorSelect.Show vbModal, sForm
GetColor = DefColor
End Function

