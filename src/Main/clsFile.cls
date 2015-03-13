VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_ContentType As String
Private m_FileName As String
Private m_Name As String
Private m_FileSize As Long
Private m_FileStart As Single
Private m_Index As Integer

Public Property Get FileIndex() As Integer
 FileIndex = m_Index
End Property

Public Property Get FormName() As String
 FormName = m_Name
End Property

Public Property Get FileName() As String
 FileName = m_FileName
End Property

Public Property Get FileExt() As String
 FileExt = GetFileExt(m_FileName)
End Property

Public Property Get FileSize() As Long
 FileSize = m_FileSize
End Property

Public Property Get ContentType() As String
 ContentType = m_ContentType
End Property

Public Property Get FileData()
 FileData = MidB(file_data, m_FileStart, m_FileSize)
End Property

Public Sub SaveAs(strFileName, Optional OverWrite = False)
 Dim f As String
 Dim fDir As String
 Dim a As Integer
 Dim fID As Integer
 Dim StrTemp As Variant
 Dim b As Long
 f = CStr(Trim(strFileName))
 f = Replace(f, "/", "\")
 a = InStr(4, f, "\", vbBinaryCompare)
 If a > 0 Then
  fDir = Left(f, a)
  Do While a > 0
   If Dir(fDir, vbDirectory) = "" Then MkDir fDir
   a = InStr(a + 1, f, "\", vbBinaryCompare)
   fDir = Left(f, a)
  Loop
 End If
 a = Len(Dir(f))
 If a > 0 Then
  If CBool(OverWrite) Then
   Kill f
  Else
   error_Number = 8
   error_Description = "该文件已经存在，不可覆盖"
   error_Source = "MyASPUploader.clsFile::SaveAs()"
   Exit Sub
  End If
 End If
 StrTemp = MidB(file_data, m_FileStart, m_FileSize)
 fID = FreeFile
 Open f For Binary As fID
  For b = 1 To LenB(StrTemp)
   Put fID, , AscB(MidB(StrTemp, b, 1))
  Next
 Close fID
End Sub

Public Sub GetFileIndex(n As Integer)
 Dim nn As Integer
 If n = m_Index And m_FileSize > 0 Then Exit Sub
 m_ContentType = CStr(file_Types.Item(n))
 m_Name = CStr(file_Names.Item(n))
 m_FileName = CStr(file_Paths.Item(n))
 m_FileSize = CLng(file_Sizes.Item(n))
 m_FileStart = 1
 For nn = 1 To n - 1
  m_FileStart = m_FileStart + file_Sizes.Item(nn)
 Next
 m_Index = n
End Sub

