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
'Private m_Index As Integer
Private m_FileData() As Byte ' As Variant

Friend Sub SetFileObj(fName, fFName, fCType, fData)
 m_Name = CStr(fName)
 m_FileName = Replace(CStr(fFName), "/", "\")
 m_ContentType = CStr(fCType)
 m_FileSize = LenB(fData)
 ReDim m_FileData(m_FileSize) As Byte
 m_FileData = fData
End Sub

'Public Property Get FileIndex() As Integer
 'FileIndex =
'End Property

Public Property Get FormName() As String
 FormName = m_Name
End Property

Public Property Get FilePath() As String
 FilePath = Left(m_FileName, InStrRev(m_FileName, "\"))
End Property

Public Property Get FileName() As String
 FileName = Mid(m_FileName, InStrRev(m_FileName, "\") + 1)
End Property

Public Property Get FileExt() As String
 FileExt = GetFileExt(m_FileName)
End Property

Public Property Get FileSize() As Long
 'FileSize = CLng(LenB(m_FileData))
 FileSize = UBound(m_FileData) + 1
End Property

Public Property Get ContentType() As String
 ContentType = m_ContentType
End Property

Public Property Get FileData()
 FileData = m_FileData
End Property

Public Sub SaveAs(strDirName, strFileName, Optional OverWrite = False)
 Dim c As String
 Dim ff As Integer
 c = CheckFileName(strDirName, strFileName)
 If c = "e" Then
  If CBool(OverWrite) Then
   Kill strDirName & strFileName
  Else
   error_Number = 6
   error_Source = "MyASPUploader.clsFile::SaveAs()"
   error_Description = "该文件已经存在"
   Err.Raise error_Number, error_Source, error_Description
   Exit Sub
  End If
 ElseIf c = "w" Then
  error_Number = 7
  error_Source = "MyASPUploader.clsFile::SaveAs()"
  error_Description = "非法的文件名“" & strDirName & strFileName & "”"
  Err.Raise error_Number, error_Source, error_Description
  Exit Sub
 End If
 MakeDirs CStr(strDirName)
 ff = FreeFile
 Open CStr(strDirName & strFileName) For Binary Access Write As #ff
 Put #ff, , m_FileData
 Close #ff
 
 'WriteData m_FileData, CStr(strDirName & strFileName)
 
End Sub
