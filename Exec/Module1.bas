Attribute VB_Name = "Module1"
Public DefColor As Single
Function GetColor(sForm As Form, Optional ByVal sDef As Single) As Single
DefColor = sDef
ColorSelect.Show vbModal, sForm
GetColor = DefColor
End Function
