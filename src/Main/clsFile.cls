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
'Private m_FileSize As Long
Private m_FileData As Variant
'Private m_Index As Integer

Friend Sub SetFileObj(fName, fFName, fCType, fData)
 m_Name = CStr(fName)
 m_FileName = CStr(fFName)
 m_ContentType = CStr(fCType)
 m_FileData = fData
End Sub

'Public Property Get FileIndex() As Integer
 'FileIndex =
'End Property

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
 FileSize = CLng(LenB(m_FileData))
End Property

Public Property Get ContentType() As String
 ContentType = m_ContentType
End Property

Public Property Get FileData()
 FileData = m_FileData
End Property

Public Sub SaveAs(strFileName, Optional OverWrite = False)
 Dim f As String
 f = CStr(Trim(strFileName))
 f = Replace(f, "/", "\")
 
 MakeDirs f
 
 If Dir(f, vbNormal + vbReadOnly + vbHidden + vbSystem) <> "" Then
  If CBool(OverWrite) Then
   Kill f
  Else
   error_Number = 6
   error_Source = "MyASPUploader.clsFile::SaveAs()"
   error_Description = "该文件已经存在"
   Err.Raise error_Number, error_Source, error_Description
   Exit Sub
  End If
 End If

 WriteData m_FileData, f
 
End Sub
