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
'Private aspScriptingContext As ScriptingContext
'Private aspResponse As Response
'Public Sub OnStartPage(PassedScriptingContext As ScriptingContext)
' Set aspScriptingContext = PassedScriptingContext
' Set aspResponse = aspScriptingContext.Response
'End Sub

Public Sub GetUpload(ByVal Data, Optional ByVal MaxSize = 0, Optional ByVal TypeLimit = "")
 Dim DivideLine, StrTemp1, StrTemp2, StrTemp3, StrTemp4, Str1, Str2
 Dim lStr1 As Long, lStr2 As Long
 Dim a As Long, b As Long, c As Long, d As Long
 Dim DataSize As Long
 Dim nName As String
 Dim nFilePath As String
 Dim nFileType As String
 
 Str1 = Str2Bin(vbCrLf)
 Str2 = Str2Bin("--")
 lStr1 = LenB(Str1)
 lStr2 = LenB(Str2)
 DataSize = LenB(Data)
 a = ConvertToByte(CStr(MaxSize))
 If DataSize > a And a > 256 Then
  Data = LeftB(Data, a)
  DataSize = a
  error_Number = 4
  error_Description = "超过限定的字节"
  error_Source = "MyASPUploader.Upload::GetUpload()"
 End If
 DivideLine = Str1 & LeftB(Data, InStrB(1, Data, Str1, vbBinaryCompare) - 1) & Str1
 Data = Str1 & LeftB(Data, DataSize - lStr1 - lStr2) & Str1
 c = LenB(DivideLine)
 a = c
 b = InStrB(a + 1, Data, DivideLine, vbBinaryCompare) - 1 - a
 Do While b > 0
  StrTemp1 = MidB(Data, a + 1, b)
  StrTemp2 = Bin2Str(LeftB(StrTemp1, InStrB(1, StrTemp1, Str1 & Str1, vbBinaryCompare) - 1))
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
  If nFilePath = "" Then
   If StrTemp1 <> "" Then AddFormObj nName, CStr(Bin2Str(StrTemp1))
  ElseIf IsTypeAllowed(GetFileExt(nFilePath), CStr(TypeLimit)) Then
   If StrTemp1 <> "" Then AddFileObj nName, nFilePath, nFileType, StrTemp1
  Else
   error_Number = 5
   error_Description = "存在限定之外的文件类型"
   error_Source = "MyASPUploader.Upload::GetUpload()"
  End If
  nName = ""
  nFilePath = ""
  a = a + b + c
  b = InStrB(a + 1, Data, DivideLine, vbBinaryCompare) - 1 - a
 Loop
 If intFileCount = 0 And error_Number = 0 Then
  error_Number = 3
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

Public Property Get GetFileByIndex(ByVal n) As clsFile
 Set GetFileByIndex = FileArray(n)
End Property

Public Property Get GetFormByIndex(ByVal n) As clsForm
 Set GetFormByIndex = FormArray(n)
End Property

Public Property Get GetFileByName(ByVal s) As clsFile
 Dim i As Integer
 For i = 1 To intFileCount
  If UCase(s) = UCase(FileArray(i).FormName) Then
   Set GetFileByName = FileArray(i)
   Exit Property
  End If
 Next
 error_Number = 1
 error_Source = "MyASPUploader.Upload::GetFileByName()"
 error_Description = "找不到此文件，无法获取"
 Err.Raise error_Number, error_Source, error_Description
End Property

Public Property Get GetFormByName(ByVal s) As clsForm
 Dim i As Integer
 For i = 1 To intFormCount
  If UCase(s) = UCase(FormArray(i).FormName) Then
   Set GetFormByName = FormArray(i)
   Exit Property
  End If
 Next
 error_Number = 2
 error_Source = "MyASPUploader.Upload::GetFormByName()"
 error_Description = "找不到此表单，无法获取"
 Err.Raise error_Number, error_Source, error_Description
End Property

Public Property Get FileCount() As Integer
 FileCount = intFileCount
End Property

Public Property Get FormCount() As Integer
 FormCount = intFormCount
End Property

Private Sub Class_Initialize()
 intFileCount = 0
 intFormCount = 0
 error_Number = 0
 error_Description = ""
 error_Source = ""
End Sub

'Private Sub Class_Terminate()

'End Sub
