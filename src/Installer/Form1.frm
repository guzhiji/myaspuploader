VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MyASPUploader Ver1.4 Installer"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3420
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   3420
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��װ"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Installer_Mode As Integer

Private Sub Command1_Click()
 On Error Resume Next
 Dim exe_dir(4) As String
 Dim exepath As String
 Dim dllpath As String
 Dim syspath As String
 Dim a As Integer
 syspath = "C:\WINDOWS"
 If Len(Dir(syspath & "\SYSTEM", vbDirectory)) = 0 Then
  syspath = "C:\WINNT"
  If Len(Dir(syspath & "\SYSTEM", vbDirectory)) = 0 Then
   MsgBox "�Ҳ���ϵͳĿ¼", vbOKOnly + vbCritical, "����"
   End
  End If
 End If
 exe_dir(0) = App.Path & "\"
 exe_dir(1) = syspath & "\"
 exe_dir(2) = syspath & "\SYSTEM32\"
 exe_dir(3) = syspath & "\SYSTEM\inetsrv\"
 exe_dir(4) = syspath & "\SYSTEM\"
 For a = 0 To 4
  If Len(Dir(exe_dir(a) & "regsvr32.exe")) > 0 Then
   exepath = exe_dir(a) & "regsvr32.exe"
   Exit For
  End If
 Next
 If Len(exepath) = 0 Then
  MsgBox "�Ҳ���ע�����regsvr32.exe����", vbOKOnly + vbCritical, "����"
  End
 End If
 Select Case Installer_Mode
  Case 0
   If Dir(App.Path & "\MyASPUploader.dll") = "" Then
    MsgBox "�Ҳ����������MyASPUploader.dll����", vbOKOnly + vbCritical, "����"
    End
   End If
   FileCopy App.Path & "\MyASPUploader.dll", exe_dir(4) & "MyASPUploader.dll"
   Shell exepath & " /s " & Chr(34) & exe_dir(4) & "MyASPUploader.dll" & Chr(34)
   If Err Then
    Err.Clear
    MsgBox "����δ֪���󣬰�װʧ�ܣ�", vbOKOnly + vbCritical, "����"
    Kill exe_dir(4) & "MyASPUploader.dll"
   Else
    MsgBox "��װ�ɹ���", vbOKOnly + vbInformation, "���"
   End If
  Case 1
   For a = 4 To 1 Step -1
    If Len(Dir(exe_dir(a) & "MyASPUploader.dll")) > 0 Then
     dllpath = exe_dir(a) & "MyASPUploader.dll"
     Exit For
    End If
   Next
   If Len(dllpath) = 0 Then
    MsgBox "�Ҳ����������MyASPUploader.dll����", vbOKOnly + vbCritical, "����"
    End
   End If
   Shell exepath & " /s /u " & Chr(34) & dllpath & Chr(34)
   If Err Then
    Err.Clear
    MsgBox "����δ֪����ж��ʧ�ܣ�", vbOKOnly + vbCritical, "����"
   Else
    Kill dllpath
    If Err Then
     Err.Clear
     MsgBox "ж�سɹ������޷�ɾ�������ļ���", vbOKOnly + vbCritical, "����"
    Else
     MsgBox "ж�سɹ���", vbOKOnly + vbInformation, "���"
    End If
   End If
 End Select
 End
End Sub

Private Sub Command2_Click()
 End
End Sub

Private Sub Form_Load()
 If IsInstalled = True Then
  Installer_Mode = 1
  Label1.Caption = "���ķ������Ѱ�װ��������Ƿ�ж�أ�"
  Command1.Caption = "ж��"
 Else
  Installer_Mode = 0
  Label1.Caption = "���ķ�����δ��װ��������Ƿ�װ��"
  Command1.Caption = "��װ"
 End If
End Sub

Private Function IsInstalled() As Boolean
 On Error Resume Next
 Dim Obj As Object
 Set Obj = CreateObject("MyASPUploader.Upload")
 If Err Then
  IsInstalled = False
  Err.Clear
 Else
  IsInstalled = True
 End If
 Set Obj = Nothing
End Function
