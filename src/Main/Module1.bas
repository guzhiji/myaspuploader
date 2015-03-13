Attribute VB_Name = "Module1"
Public FileArray() As clsFile
Public FormArray() As clsForm
Public intFileCount As Integer
Public intFormCount As Integer
'Public file_Names As Collection
'Public file_Paths As Collection
'Public file_Types As Collection
'Public file_Sizes As Collection
'Public form_Names As Collection
'Public form_Values As Collection
'Public file_data As Variant
Public error_Number As Integer
Public error_Description As String
Public error_Source As String

Public Function GetFileExt(ByVal uFileName As String) As String
 GetFileExt = Trim(Right(uFileName, Len(uFileName) - InStrRev(uFileName, ".")))
End Function

Public Sub AddFileObj(fName, fFName, fCType, fData)
 intFileCount = intFileCount + 1
 ReDim Preserve FileArray(1 To intFileCount)
 Set FileArray(intFileCount) = New clsFile
 FileArray(intFileCount).SetFileObj fName, fFName, fCType, fData
End Sub

Public Sub AddFormObj(fName, fValue)
 intFormCount = intFormCount + 1
 ReDim Preserve FormArray(1 To intFormCount)
 Set FormArray(intFormCount) = New clsForm
 FormArray(intFormCount).SetFormObj fName, fValue
End Sub

Public Sub MakeDirs(ByVal strPathName As String)
 Dim fDir As String
 Dim a As Integer
 a = InStr(4, strPathName, "\", vbBinaryCompare)
 If a > 0 Then
  fDir = Left(strPathName, a)
  Do While a > 0
   If Dir(fDir, vbDirectory) = "" Then MkDir fDir
   a = InStr(a + 1, strPathName, "\", vbBinaryCompare)
   fDir = Left(strPathName, a)
  Loop
 End If
End Sub

Public Sub WriteData(ByVal fData, ByVal fName As String)
 Dim fID As Integer
 Dim p As Long, l As Long
 l = LenB(fData)
 fID = FreeFile
 Open fName For Binary As fID
  For p = 1 To l
   Put fID, , AscB(MidB(fData, p, 1))
  Next
 Close fID
End Sub

Public Function ConvertToByte(strSize As String) As Long
 Dim a As String
 ConvertToByte = 0
 strSize = LCase(Trim(strSize))
 If IsNumeric(strSize) Then
  ConvertToByte = CLng(strSize)
 ElseIf Len(strSize) > 2 Then
  a = Left(strSize, Len(strSize) - 2)
  Select Case Right(strSize, 2)
   Case "kb"
    If IsNumeric(a) Then ConvertToByte = CLng(CSng(a) * 1024)
   Case "mb"
    If IsNumeric(a) Then ConvertToByte = CLng(CSng(a) * 1024 ^ 2)
   Case "gb"
    If IsNumeric(a) Then ConvertToByte = CLng(CSng(a) * 1024 ^ 3)
   Case Else
    If Len(strSize) > 4 Then
     If Right(strSize, 4) = "byte" Then
      a = Left(strSize, Len(strSize) - 4)
      If IsNumeric(a) Then ConvertToByte = CLng(a)
     End If
    End If
  End Select
 End If
End Function

Public Function IsTypeAllowed(TheType As String, AllTypes As String) As Boolean
 Dim a, b
 Dim aTypes As String
 aTypes = Trim(LCase(AllTypes))
 If Len(aTypes) > 0 Then
  aTypes = Replace(aTypes, ",", ";")
  aTypes = Replace(aTypes, "|", ";")
  aTypes = Replace(aTypes, "\", ";")
  aTypes = Replace(aTypes, "/", ";")
  a = Split(aTypes, ";")
  IsTypeAllowed = False
  For Each b In a
   If Len(b) > 0 Then
    If CStr(b) = LCase(TheType) Then
     IsTypeAllowed = True
     Exit Function
    End If
   End If
  Next
 Else
  IsTypeAllowed = True
 End If
End Function

Public Function URLDecode(Expression) As String
 On Error Resume Next
 Dim s, r
 Dim hCode, hTemp
 Dim iCode, iL, i
 hTemp = ""
 r = ""
 s = Replace(Expression, "+", " ")
 s = Split(s, "%")
 For i = LBound(s) + 1 To UBound(s)
  iL = Len(s(i))
  If iL >= 2 Then
   hCode = Left(s(i), 2)
   iCode = CLng("&H" & hCode)
   If Not Err Then
    If hTemp = "" Then
     If iCode > 127 And iL = 2 Then
      hTemp = hCode
     Else
      r = r & Chr(iCode)
     End If
    Else
     hCode = hTemp & hCode
     hTemp = ""
     iCode = CLng("&H" & hCode)
     r = r & Chr(iCode)
    End If
    If iL > 2 Then r = r & Right(s(i), iL - 2)
   Else
    Err.Clear
    If hTemp <> "" Then
     r = r & hTemp
     hTemp = ""
    End If
    r = r & s(i)
   End If
  Else
   If hTemp <> "" Then
    r = r & hTemp
    hTemp = ""
   End If
   r = r & s(i)
  End If
 Next
 r = s(LBound(s)) & r
 URLDecode = CStr(r)
End Function

Public Function Str2Bin(UStr)
 Dim p As Long, l As Long
 Dim c
 Str2Bin = ""
 If IsNull(UStr) Then Exit Function
 l = Len(UStr)
 For p = 1 To l
  c = Mid(UStr, p, 1)
  Str2Bin = Str2Bin & ChrB(AscB(c))
 Next
End Function

Public Function Bin2Str(ByVal BStr)
 Dim i As Long, l As Long
 Dim SkipFlag As Integer
 Dim r, t
 
 SkipFlag = 0
 r = ""

 If Not IsNull(BStr) Then
  l = LenB(BStr)
  For i = 1 To l
   If SkipFlag = 0 Then
    t = MidB(BStr, i, 1)
    If AscB(t) > 127 And i < l Then
     r = r & Chr(AscW(MidB(BStr, i + 1, 1) & t))
     SkipFlag = 1
    Else
     r = r & Chr(AscB(t))
    End If
   Else
    SkipFlag = 0
   End If
  Next
 End If
 Bin2Str = r
End Function

