Attribute VB_Name = "Module1"
Public file_Names As Collection
Public file_Temps As Collection
Public file_Paths As Collection
Public file_Types As Collection
Public file_Sizes As Collection
Public form_Names As Collection
Public form_Values As Collection
Public error_Number As Integer
Public error_Description As String
Public error_Source As String

Public Function GetFileExt(ByVal uFileName As String) As String
 GetFileExt = Trim(Right(uFileName, Len(uFileName) - InStrRev(uFileName, ".")))
End Function
