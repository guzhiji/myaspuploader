VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Upload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private uFile As clsFile
Private uForm As clsForm

'Private aspScriptingContext As ScriptingContext
'Private aspResponse As Response
'Public Sub OnStartPage(PassedScriptingContext As ScriptingContext)
' Set aspScriptingContext = PassedScriptingContext
' Set aspResponse = aspScriptingContext.Response
'End Sub

Public Sub GetUpload(ByVal Data, Optional ByVal MaxSize = 0, Optional ByVal TypeLimit = "")
 Dim DivideLine, StrTemp1, StrTemp2, StrTemp3, StrTemp4, Str1, Str2
 Dim lStr1 As Long
 Dim lStr2 As Long
 Dim DataSize As Long
 Dim a As Long
 Dim b As Long
 Dim c As Long
 Dim d As Long
 Dim fID As Integer
 Dim nName As String
 Dim nFileTemp As String
 Dim nFilePath As String
 Dim nFileType As String
 Dim nFileSize As Long
 
 Str1 = UStr2BStr(vbCrLf)
 Str2 = UStr2BStr("--")
 lStr1 = LenB(Str1)
 lStr2 = LenB(Str2)
 DataSize = LenB(Data)
 a = ConvertToByte(CStr(MaxSize))
 If DataSize > a And a > 256 Then
  Data = LeftB(Data, a)
  DataSize = a
  error_Number = 6
  error_Description = "超过限定的 " & DataSize & " 字节"
  error_Source = "MyASPUploader.Upload::GetUpload()"
 End If
 DivideLine = Str1 & LeftB(Data, InStrB(1, Data, Str1, vbBinaryCompare) - 1) & Str1
 Data = Str1 & LeftB(Data, DataSize - lStr1 - lStr2) & Str1
 c = LenB(DivideLine)
 a = c
 b = InStrB(a + 1, Data, DivideLine, vbBinaryCompare) - 1 - a
 Do While b > 0
  StrTemp1 = MidB(Data, a + 1, b)
  StrTemp2 = BStr2UStr(LeftB(StrTemp1, InStrB(1, StrTemp1, Str1 & Str1, vbBinaryCompare) - 1))
  StrTemp1 = RightB(StrTemp1, LenB(StrTemp1) - InStrB(1, StrTemp1, Str1 & Str1, vbBinaryCompare) - lStr1 - lStr1 + 1)
  StrTemp2 = Split(StrTemp2, vbCrLf)
  For Each StrTemp3 In StrTemp2
   d = InStr(1, StrTemp3, "Content-Type: ", vbBinaryCompare)
   If d > 0 Then
    nFileType = CStr(Trim(Right(StrTemp3, Len(StrTemp3) - 14)))
   ElseIf Len(StrTemp3) > 0 Then
    StrTemp3 = Split(StrTemp3, ";")
    For Each StrTemp4 In StrTemp3
     If Len(nName) = 0 Then
      d = InStr(1, StrTemp4, "name=" & Chr(34), vbBinaryCompare)
      If d > 0 Then nName = CStr(Trim(Right(StrTemp4, Len(StrTemp4) - d - 5))): nName = Left(nName, Len(nName) - 1)
     End If
     If Len(nFilePath) = 0 Then
      d = InStr(1, StrTemp4, "filename=" & Chr(34), vbBinaryCompare)
      If d > 0 Then nFilePath = CStr(Trim(Right(StrTemp4, Len(StrTemp4) - d - 9))): nFilePath = URLDecode(Left(nFilePath, Len(nFilePath) - 1))
     End If
    Next
   End If
  Next
  If Len(nFilePath) = 0 Then
   nFileSize = CLng(LenB(StrTemp1))
   If nFileSize > 0 Then
    form_Names.Add nName
    form_Values.Add CStr(BStr2UStr(StrTemp1))
   End If
  ElseIf IsTypeAllowed(GetFileExt(nFilePath), CStr(TypeLimit)) Then
   nFileSize = CLng(LenB(StrTemp1))
   If nFileSize > 0 Then
    nFileTemp = NewTempFile
    file_Temps.Add nFileTemp
    file_Sizes.Add nFileSize
    file_Types.Add nFileType
    file_Paths.Add nFilePath
    file_Names.Add nName
    
    fID = FreeFile
    Open nFileTemp For Binary As fID
     For d = 1 To LenB(StrTemp1)
      Put fID, , AscB(MidB(StrTemp1, d, 1))
     Next
    Close fID
   End If
  Else
   error_Number = 7
   error_Description = "存在限定之外的文件类型"
   error_Source = "MyASPUploader.Upload::GetUpload()"
  End If
  nName = ""
  nFilePath = ""
  a = a + b + c
  b = InStrB(a + 1, Data, DivideLine, vbBinaryCompare) - 1 - a
 Loop
 If file_Names.Count = 0 And error_Number = 0 Then
  error_Number = 5
  error_Description = "文件个数为0"
  error_Source = "MyASPUploader.Upload::GetUpload()"
 End If
End Sub

Public Property Get ErrorNumber() As Integer
 ErrorNumber = error_Number
End Property

Public Property Get ErrorDescription() As String
 ErrorDescription = error_Description
End Property

Public Property Get ErrorSource() As String
 ErrorSource = error_Source
End Property

Public Property Get GetFile(ByVal n, Optional ByVal mode = 0) As clsFile
 Dim a As Integer
 Dim b As Integer
 b = file_Names.Count
 If b = 0 Then Err.Raise 1, "MyASPUploader.Upload::GetFile()", "文件个数为0，无法获取"
 Select Case CInt(mode)
  Case 0
   For a = 1 To b
    If LCase(n) = LCase(file_Names.Item(a)) Then Exit For
   Next
  Case Else
   a = CInt(n) + 1
 End Select
 If a >= 1 And a <= b Then
  If uFile Is Nothing Then Set uFile = New clsFile
  uFile.GetFileIndex a
  Set GetFile = uFile
 Else
  Err.Raise 2, "MyASPUploader.Upload::GetFile()", "找不到此文件，无法获取"
 End If
End Property

Public Property Get GetForm(ByVal n, Optional ByVal mode = 0) As clsForm
 Dim a As Integer
 Dim b As Integer
 b = form_Names.Count
 If b = 0 Then Err.Raise 3, "MyASPUploader.Upload::GetForm()", "表单个数为0，无法获取"
 Select Case CInt(mode)
  Case 0
   For a = 1 To b
    If LCase(n) = LCase(CStr(form_Names.Item(a))) Then Exit For
   Next
  Case Else
   a = CInt(n) + 1
 End Select
 If a >= 1 And a <= b Then
  If uForm Is Nothing Then Set uForm = New clsForm
  uForm.GetFormIndex a
  Set GetForm = uForm
 Else
  Err.Raise 4, "MyASPUploader.Upload::GetForm()", "找不到此表单，无法获取"
 End If
End Property

Public Property Get FileCount() As Integer
 FileCount = file_Names.Count
End Property

Public Property Get FormCount() As Integer
 FormCount = form_Names.Count
End Property

Private Function ConvertToByte(strSize As String) As Long
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

Private Function IsTypeAllowed(TheType As String, AllTypes As String) As Boolean
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

Private Function URLDecode(Expression) As String
 Dim strSource, StrTemp, strResult
 Dim lngPos
 strSource = Replace(Expression, "+", " ")
 For lngPos = 1 To Len(strSource)
  StrTemp = Mid(strSource, lngPos, 1)
  If StrTemp = "%" Then
   If lngPos + 2 < Len(strSource) Then
    strResult = strResult & Chr(CInt("&H" & Mid(strSource, lngPos + 1, 2)))
    lngPos = lngPos + 2
   End If
  Else
   strResult = strResult & StrTemp
  End If
 Next
 URLDecode = CStr(strResult)
End Function

Private Function NewTempFile() As String
 If Len(Dir(App.Path & "\temp", vbDirectory)) = 0 Then MkDir App.Path & "\temp"
 Dim a As Integer
 a = 0
 Do While Len(Dir(App.Path & "\temp\" & CStr(a) & ".tmp")) > 0
  a = a + 1
 Loop
 NewTempFile = App.Path & "\temp\" & CStr(a) & ".tmp"
End Function

Private Function UStr2BStr(UStr)
 Dim lngLoop
 Dim strChar
 UStr2BStr = ""
 For lngLoop = 1 To Len(UStr)
  strChar = Mid(UStr, lngLoop, 1)
  UStr2BStr = UStr2BStr & ChrB(AscB(strChar))
 Next
End Function

Private Function BStr2UStr(BStr)
 Dim lngLoop
 BStr2UStr = ""
 BStr2UStr = BtoS(BStr)
End Function

Private Function BtoS(varstr)
 Dim str2bin
 Dim varchar
 Dim varasc
 Dim varlow, varhigh
 Dim i
 str2bin = ""
 ' 测试
 skipflag = 0
 strC = ""
 binstr = varstr
 If Not IsNull(binstr) Then
  lnglen = LenB(binstr)
  For i = 1 To lnglen
   If skipflag = 0 Then
    tmpBin = MidB(binstr, i, 1)
    '判断是否中文的字符
    If AscB(tmpBin) > 127 Then
     'AscW会把二进制的中文双字节字符高位和低位反转，所以要先把中文的高低位反转
     strC = strC & Chr(AscW(MidB(binstr, i + 1, 1) & tmpBin))
     skipflag = 1
    Else
     strC = strC & Chr(AscB(tmpBin))
    End If
   Else
    skipflag = 0
   End If
  Next
 End If
 BtoS = strC
End Function

Private Sub Class_Initialize()
 Set file_Names = New Collection
 Set file_Temps = New Collection
 Set file_Paths = New Collection
 Set file_Types = New Collection
 Set file_Sizes = New Collection
 Set form_Names = New Collection
 Set form_Values = New Collection
 error_Number = 0
 error_Description = ""
 error_Source = ""
End Sub

Private Sub Class_Terminate()
 Dim a As Integer
 Dim b As Integer
 a = file_Temps.Count
 For b = 1 To a
  Kill CStr(file_Temps.Item(b))
 Next
 
 Set file_Names = Nothing
 Set file_Temps = Nothing
 Set file_Paths = Nothing
 Set file_Types = Nothing
 Set file_Sizes = Nothing
 Set form_Names = Nothing
 Set form_Values = Nothing
End Sub
