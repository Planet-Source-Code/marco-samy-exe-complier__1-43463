Attribute VB_Name = "MCO"
'Color Controller Module
'Copyright (c) Marco Samy 2002
Function ColorRGB(ByVal sColor As Single, ByRef ColorR, ByRef ColorG, ByRef ColorB)
    ColorR = (sColor And 255)
    ColorG = (sColor And 65280) / 256
    ColorB = (sColor And 16711680) / 65536
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
Function RndColor()
Dim R, G, B
Randomize
R = Int(255 * Rnd) + 1
G = Int(255 * Rnd) + 1
B = Int(255 * Rnd) + 1
RndColor = RGB(R, G, B)
End Function
