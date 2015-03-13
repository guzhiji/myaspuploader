VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_Name As String
Private m_Value As String
'Private m_Index As Integer

Friend Sub SetFormObj(fName, fValue)
 m_Name = CStr(fName)
 m_Value = CStr(fValue)
End Sub

'Public Property Get FormIndex() As Integer
' FormIndex = m_Index
'End Property

Public Property Get FormName() As String
 FormName = m_Name
End Property

Public Property Get FormValue() As String
 FormValue = m_Value
End Property
