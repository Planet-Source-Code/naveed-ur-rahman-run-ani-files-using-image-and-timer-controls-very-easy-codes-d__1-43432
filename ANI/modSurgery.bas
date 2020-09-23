Attribute VB_Name = "modSurgery"
'This module is a part of modSurgery
'Written by Naveed Ur Rahman (neenojee@hotmail.com)

'This module give excellent functions for handling
'the binary of files.

'I am not describing the functions.
'The names of the FUNCTIONS are self-descriptive,
'aren't they ?

'--------------------------------------------------
Sub WriteFile(ReadFile, WriteFile, FromByte, Size)
Dim ReadString  As String

Dim FreeFR, FreeFW

FreeFR = FreeFile
Open ReadFile For Binary As #FreeFR
FreeFW = FreeFile
Open WriteFile For Binary As #FreeFW

If Size < 2048 Then

ReadString = Space(Size)
Get #FreeFR, FromByte, ReadString
Put #FreeFW, LOF(FreeFW) + 1, ReadString

Else

a = Size / 2048
b = Size \ 2048
a = (a - b)

ReadString = Space(2048)

For z = FromByte To b * 2048 + FromByte Step 2048
Get #FreeFR, z, ReadString
Put #FreeFW, LOF(FreeFW) + 1, ReadString
Next z

If a > 0 Then
ReadString = Space(Round(a * 2048, 0))
Get #FreeFR, (b * 2048 + FromByte) + 1, ReadString
Put #FreeFW, LOF(FreeFW) + 1, ReadString
End If

End If

Close #FreeFR
Close #FreeFW
End Sub


Function GetString(ReadFile, FromByte, Size)
Dim FreeF, ReadString As String
ReadString = Space(Size)
FreeF = FreeFile
Open ReadFile For Binary As #FreeF
Get #FreeF, FromByte, ReadString
Close #FreeF
GetString = ReadString
ReadString = ""
End Function


Function PutString(WriteFile, FromByte, String2Write)
Dim WriteString As String
WriteString = String2Write
Dim FreeF
FreeF = FreeFile
Open WriteFile For Binary As #FreeF
Put #FreeF, FromByte, WriteString
Close #FreeF
End Function

Function GetIntMultiSize(Optional N1 As Long, Optional N2 As Long, Optional N3 As Long, Optional N4 As Long) As Long
GetIntMultiSize = N1 + N2 * 256 + N3 * 256 ^ 2 + N4 * 256 ^ 3
End Function

Sub Advertise()
'Do you have 'Naveed IconEX 5.00 (Second Edition - XP-Look)' ?
'If no then "Download It Now For The Free !!!"
'Website: www.iconex.0catch.com

Dim AnsMsg As VbMsgBoxResult
AnsMsg = MsgBox("Do you have 'Naveed IconEX 5.00 (Second Edition - XP-Look)' ?", vbYesNo + vbQuestion + vbSystemModal, "Naveed IconEX 5.00 (Second Edition - XP-Look)")
If AnsMsg = vbNo Then
MsgBox "Download It Now For The Free !!!" & vbCrLf & "www.iconex.0catch.com", vbInformation, "Download"
Shell "start www.iconex.0catch.com", vbHide
End If
End Sub
'--------------------------------------------------

