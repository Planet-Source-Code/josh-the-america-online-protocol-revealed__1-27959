Attribute VB_Name = "Module1"
Function AOLCRC(strng$, lenstr)
Dim crc As Long
Dim ch As Long
Dim i As Long
Dim j As Long

For i = 0 To lenstr - 1
ch = Asc(Mid(strng$, i + 1, 1))
For j = 0 To 7
If ((crc Xor ch) And 1) Then
crc = (Int((crc / 2)) Xor 40961)
Else
crc = Int(crc / 2)
End If
ch = Int(ch / 2)
Next j
Next i
AOLCRC = crc
End Function

Function give5th(pack As String, inputz As Integer)
'You need the EnHex function for this, you will find it later on in this document.
Dim give5th1 As Integer
inputz = inputz + 4
give5th1 = Len(DeHex(pack)) - inputz
give5th = EnHex(Chr(give5th1))
End Function

Function gethibyte(mybyte As Variant) As Variant
gethibyte = Int(mybyte / 256)
End Function

Function getlobyte(mybyte As Variant) As Variant
getlobyte = Int((mybyte - (gethibyte(mybyte) * 256)))
End Function


Function stripnulls(thedata As String)
Dim torem As Integer
torem = 1
Do Until crap = 0
torem = InStr(1, thedata, Chr$(0))
If torem > 0 Then Mid(thedata, torem, 1) = "Å"
Loop
End Function
Public Function EnHex(Data As String) As String
    Dim iCount As Double
    Dim sTemp As String


    For iCount = 1 To Len(Data)
        sTemp = Hex$(Asc(Mid$(Data, iCount, 1)))
        If Len(sTemp) < 2 Then sTemp = "0" & sTemp
        EnHex = EnHex & sTemp
    Next iCount
End Function

Public Function DeHex(Data As String) As String
    Dim iCount As Double


    For iCount = 1 To Len(Data) Step 2
        DeHex = DeHex & Chr$(Val("&H" & Mid$(Data, iCount, 2)))
    Next iCount
End Function

Public Function Int2Hex(Data As String) As String
    Dim sTemp As String
        sTemp = Hex(Data)
        If Len(sTemp) < 2 Then sTemp = "0" & sTemp
        Int2Hex = Int2Hex & sTemp
  
End Function

Public Function Hex2Int(Data As String) As String
    Dim iCount As Double
    For iCount = 1 To Len(Data) Step 2
        Hex2Int = Hex2Int & CInt(Val("&H" & Mid$(Data, iCount, 2)))
    Next iCount
End Function


