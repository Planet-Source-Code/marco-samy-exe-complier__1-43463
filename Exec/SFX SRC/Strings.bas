Attribute VB_Name = "Strings"
'/////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////
'///////////////Strings Control Module For Presentation///////////////////////////
'///////////////Programmed By:  Marco Samy Nasif   1999/////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////
Public Function FindFinal(ByVal sti As String, ByVal Insti As String) As Integer
On Error Resume Next
FindFinal = InStrRev(Insti, sti, Len(Insti), vbTextCompare)
End Function
'---------------
Public Function DoUp(ByVal sPath As String) As String
If Right$(sPath, 1) = "\" Then sPath = Left$(sPath, Len(sPath) - 1)
Dim Xs
Xs = FindFinal("\", sPath)
If Len(sPath) = 3 Or Len(sPath) = 2 Then
DoUp = sPath & "\"
Else
If Xs = 3 Then
DoUp = Left$(sPath, Xs)
Else
DoUp = Left$(sPath, Xs - 1)
End If
End If
End Function
'------------------------Word Wraping .............
Public Function TforPic(ByVal sText As String, sPic As PictureBox, Optional ByVal Margins As Double = 10) As String
Dim IC As New Collection, FC As New Collection, LC As New Collection, I, x, MyStr As String, MyWord As String
GetAllAB sText, vbCrLf, vbCrLf, IC
If IC.Count = 0 Then IC.Add sText
    For I = 1 To IC.Count
    EmptyColl FC
    GetAllAB IC.Item(I), " ", " ", FC
    If FC.Count = 0 Then FC.Add IC.Item(I)
        While Not FC.Count = 0
        MyStr = ""
            While Val(sPic.TextWidth(MyStr)) < Val(sPic.Width - (2 * Margins) - 30)
            If Not FC.Count = 0 Then
            MyStr = MyStr & " " & FC.Item(1)
            FC.Remove (1)
            Else: GoTo PassHere
            End If
            Wend
If FindFinal(" ", MyStr) = (0 Or 1) Then GoTo PassHere
        MyWord = GetAL(" ", MyStr)
        MyStr = GetBL(" ", MyStr)
        If FC.Count = 0 Then FC.Add MyWord Else FC.Add MyWord, , 1
PassHere:
        LC.Add MyStr
        Wend
'    LC.Add ""
    Next I
TforPic = LC.Item(1)
sPic.CurrentX = Margins
sPic.Print LC.Item(1)
For I = 2 To LC.Count
TforPic = TforPic & vbCrLf & LC.Item(I)
sPic.CurrentX = Margins
sPic.Print LC.Item(I)
Next I
End Function
Function EmptyColl(sColl As Collection)
For Z = 1 To sColl.Count
sColl.Remove (1)
Next Z
End Function
Public Function TforPic2(ByVal sText As String, sPic As PictureBox, ToCol As Collection, Optional ByVal Margins As Double = 10) As String
Dim IC As New Collection, FC As New Collection, LC As New Collection, I, x, MyStr As String, MyWord As String
GetAllAB sText, vbCrLf, vbCrLf, IC
If IC.Count = 0 Then IC.Add sText
    For I = 1 To IC.Count
    EmptyColl FC
    GetAllAB IC.Item(I), " ", " ", FC
    If FC.Count = 0 Then FC.Add IC.Item(I)
        While Not FC.Count = 0
        MyStr = ""
            While Val(sPic.TextWidth(MyStr)) < Val(sPic.Width - (2 * Margins) - 30)
            If Not FC.Count = 0 Then
            MyStr = MyStr & " " & FC.Item(1)
            FC.Remove (1)
            Else: GoTo PassHere
            End If
            Wend
If FindFinal(" ", MyStr) = (0 Or 1) Then GoTo PassHere
        MyWord = GetAL(" ", MyStr)
        MyStr = GetBL(" ", MyStr)
        If FC.Count = 0 Then FC.Add MyWord Else FC.Add MyWord, , 1
PassHere:
        ToCol.Add MyStr
        Wend
    Next I
TforPic2 = ToCol.Item(1)
For I = 2 To ToCol.Count
TforPic2 = TforPic2 & vbCrLf & ToCol.Item(I)
Next I
End Function
'---------------
Public Function FindAll(ByVal sti As String, ByVal Insti As String, ByRef sColl As Collection)
Dim O0 As Integer, O1 As Integer
O0 = 1
Point1:
O1 = InStr(O0, Insti, sti)
If (O1 = 0) And (O0 = 1) Then Exit Function
If O1 > 0 Then
sColl.Add O1
O0 = O1 + 1
GoTo Point1
Else
Exit Function
End If
End Function
Public Function FindCount(ByVal sti As String, ByVal Insti As String) As Single
FindCount = 0
Dim O0 As Integer, O1 As Integer
O0 = 1
Point1:
O1 = InStr(O0, Insti, sti)
If (O1 = 0) And (O0 = 1) Then Exit Function
If O1 > 0 Then
FindCount = FindCount + 1
O0 = O1 + 1
GoTo Point1
Else
Exit Function
End If
End Function

'--------------------
Public Function RemoveFirst(ByVal sti As String, ByVal Insti As String, ByVal Bpo As Integer) As String
Dim tx1, tx2, l1, z1, S1
S1 = Len(sti)
z1 = InStr(Bpo, Insti, sti, vbTextCompare)
tx1 = Left$(Insti, z1 - 1)
l1 = Len(Insti) - (Len(tx1) + S1)
tx2 = Right$(Insti, l1)
RemoveFirst = tx1 & tx2
End Function
'------------------
Public Function PutS(ByVal sti As String, ByVal Insti As String, ByVal Apo As Integer) As String
Dim tx2, tx1, l1
tx1 = Left$(Insti, Apo)
l1 = Len(Insti) - Len(tx1)
tx2 = Right$(Insti, l1)
PutS = tx1 & sti & tx2
End Function
'----------------------
Public Function ReplaceF(ByVal Stix As String, ByVal InStix As String, ByVal repBy As String, ByVal begP As Integer) As String
ReplaceF = Replace(InStix, Stix, repBy, begP, 1, vbTextCompare)
End Function
'-----------------
Public Function ReplaceAll1(ByVal Stix As String, ByVal InStix As String, ByVal repBy As String, ByVal begP As Integer) As String
ReplaceAll = Replace(InStix, Stix, repBy, begP, -1, vbTextCompare)
End Function
'---------------------
Public Function OneDi(sPath As String) As Boolean
Dim SS As New Collection
FindAll "\", sPath, SS
If SS.Count = 1 Then OneDi = True Else OneDi = False
End Function
'---------------------------
Public Function ReplaceL(ByVal sts As String, ByVal Insts As String, ByVal repBy As String) As String
Dim oz
oz = FindFinal(sts, Insts)
ReplaceL = ReplaceF(sts, Insts, repBy, oz)
End Function
'------------------
Public Function MakeString(ByVal Stri As String, ByVal nums As Integer) As String
If nums = Null Or nums <= 0 Then
MakeString = ""
Exit Function
End If
Dim os, ov, x
os = Stri
ov = Stri
For x = 1 To nums
ov = ov & os
Next x
MakeString = ov

End Function
'--------------------
Function PutAEWord(ByVal sWard As String, ByVal swhat As String, ByVal InWhat As String) As String
Dim tex1
tex1 = ReplaceAll1(sWard, InWhat, sWard & swhat, 1)
PutAEWord = tex1
End Function
Function PutBEWord(ByVal sWard As String, ByVal swhat As String, ByVal InWhat As String) As String
Dim tex1
tex1 = ReplaceAll1(sWard, InWhat, swhat & sWard, 1)
PutBEWord = tex1
End Function
Function PutAWord(ByVal sWard As String, ByVal swhat As String, ByVal InWhat As String) As String
Dim tex1
tex1 = ReplaceF(sWard, InWhat, sWard & swhat, 1)
PutAWord = tex1
End Function
Function PutBWord(ByVal sWard As String, ByVal swhat As String, ByVal InWhat As String) As String
Dim tex1
tex1 = ReplaceF(sWard, InWhat, swhat & sWard, 1)
PutBWord = tex1
End Function
Function RemoveAll(ByVal sWard As String, ByVal InWhat As String) As String
Dim tex1
tex1 = ReplaceAll1(sWard, InWhat, vbNullString, 1)
RemoveAll = tex1
End Function
'---------------
'to match whole word only and not case
Function MatchWW(ByVal Sword As String, ByVal InWhat As String, ByVal BegPo As String)
Dim sw
sw = " " & Sword & " "
MatchWW = InStr(BegPo, InWhat, sw)
End Function
'----------------
'to get letter from string
Function GetL(ByVal sPoint As Integer, ByVal sts As String) As String
GetL = Mid$(sts, sPoint, 1)
End Function
'-----------------
Function GetfAb(ByVal Insti As String, ByVal Asti As String, ByVal BSti As String, ByRef Sp As Integer, ByRef Ep As Integer, ByVal Apo As Integer) As String
Dim FA, FB, LA, LB, Ap
Ap = Apo
Point1:
FA = InStr(Ap, Insti, Asti)
If FA = 0 Then
GetfAb = vbNullString
Exit Function
End If
LA = Len(Asti)
LB = Len(BSti)
FB = InStr(FA + LA, Insti, BSti)
If FB = 0 Then
GetfAb = vbNullString
Exit Function
End If
If Not (FB - FA) = LA Then
GetfAb = Mid$(Insti, FA + LA, FB - (FA + LA))
Sp = FA + LA
Ep = FB - 1
Else
If (Val(FindFinal(Asti, Insti)) <= Val(FA)) Then
GetfAb = vbNullString
Exit Function
Else
Ap = FB + LB
GoTo Point1
End If
End If
End Function
Function Between(ByVal Insti As String, ByVal Asti As String, ByVal BSti As String, ByVal Apo As Integer) As String
Dim FA, FB, LA, LB, Ap
Ap = Apo
Point1:
FA = InStr(Ap, Insti, Asti)
If FA = 0 Then
Between = vbNullString
Exit Function
End If
LA = Len(Asti)
LB = Len(BSti)
FB = InStr(FA + LA, Insti, BSti)
If FB = 0 Then
Between = vbNullString
Exit Function
End If
If Not (FB - FA) = LA Then
Dim TempStr
TempStr = Left$(Insti, FB)
FA = FindFinal(Asti, Insti)
Between = Mid$(Insti, FA + LA, FB - (FA + LA))
Else
If (Val(FindFinal(Asti, Insti)) <= Val(FA)) Then
Between = vbNullString
Exit Function
Else
Ap = FB + LB
GoTo Point1
End If
End If
End Function
'-------------
Function GetAL(ByVal sti As String, ByVal Insti As String) As String
Dim al
al = FindFinal(sti, Insti)
If ((al = 0) Or (al = Len(Insti) - (Len(sti) - 1))) Then
GetAL = vbNullString
Exit Function
End If
GetAL = Right$(Insti, (Len(Insti) - (al + Len(sti) - 1)))
End Function
'---------------
Function GetBF(ByVal sti As String, ByVal Insti As String, ByVal Bpo As Integer) As String
Dim Bf
Bf = InStr(Bpo, Insti, sti)
If Val(Bf) < Val(1) Then
GetBF = vbNullString
Exit Function
End If
GetBF = Left$(Insti, Bf - 1)
End Function
Function GetBL(ByVal sti As String, ByVal Insti As String) As String
Dim Bl
Bl = FindFinal(sti, Insti)
If Val(Bl) < Val(1) Then
GetBL = vbNullString
Exit Function
End If
GetBL = Left$(Insti, Bl - 1)
End Function
Function GetAF(ByVal sti As String, ByVal Insti As String, ByVal Bpo As Integer) As String
Dim af
af = InStr(Bpo, Insti, sti)
If ((af = 0) Or (af = Len(Insti) - (Len(sti) - 1))) Then
GetAF = vbNullString
Exit Function
End If
GetAF = Right$(Insti, (Len(Insti) - (af + Len(sti) - 1)))
End Function
'------------------
Function GetAllAB(ByVal sText As String, ByVal Asti As String, ByVal BSti As String, sColl As Collection)
Dim Nl, Spl As Integer, EPl As Integer, FW, LW
FW = GetBF(BSti, sText, 1)
If Not FW = vbNullString Then sColl.Add FW
LW = GetAL(Asti, sText)
Nl = GetfAb(sText, Asti, BSti, Spl, EPl, 1)
While Not Nl = vbNullString
sColl.Add Nl
Nl = GetfAb(sText, Asti, BSti, Spl, EPl, EPl)
Wend
If Not LW = vbNullString Then sColl.Add LW
End Function
Function GetAllAB2(ByVal sText As String, ByVal Asti As String, ByVal BSti As String, sColl As Collection)
Dim Nl, Spl As Integer, EPl As Integer, FW, LW
FW = GetBF(BSti, sText, 1)
If Not FW = vbNullString Then sColl.Add FW
LW = GetAL(Asti, sText)
Nl = GetfAb(sText, Asti, BSti, Spl, EPl, 1)
While Not Nl = vbNullString
sColl.Add Nl
Nl = GetfAb(sText, Asti, BSti, Spl, EPl, EPl + 1)
Wend
If Not LW = vbNullString Then sColl.Add LW
End Function
'------------------


