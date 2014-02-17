VERSION 5.00
Begin VB.Form Form_rest 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Data Barang"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form_rest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Restore Database"
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.TextBox Text1 
         Height          =   510
         Left            =   1080
         TabIndex        =   2
         Top             =   840
         Width           =   3615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000C0&
         Caption         =   "Proses"
         Height          =   735
         Left            =   1560
         TabIndex        =   1
         Top             =   1680
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Form_rest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim interval As Integer
Option Explicit

Private Sub execCommand(ByVal cmd As String)
    Dim result  As Long
    Dim lPid    As Long
    Dim lHnd    As Long
    Dim lRet    As Long

    cmd = "cmd /c " & cmd
    result = Shell(cmd, vbHide)

    lPid = result
    If lPid <> 0 Then
        lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
        If lHnd <> 0 Then
            lRet = WaitForSingleObject(lHnd, INFINITE)
            CloseHandle (lHnd)
        End If
    End If
End Sub

Private Sub Command1_Click()
    Dim cmd As String
    On Error Resume Next

    cmd = Chr(34) & "D:\Xampp\MySQL\bin\mysqldump" & Chr(34) & " -uroot  --routines --comments invent_db < c:\Backup\" & Text1.Text & ""
    Call execCommand(cmd)

    MsgBox "Restore Database berhasil", vbInformation
End Sub
