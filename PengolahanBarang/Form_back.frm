VERSION 5.00
Begin VB.Form Form_back 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Data Barang"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6195
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
   Icon            =   "Form_back.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Backup Database"
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5535
      Begin VB.TextBox txttgl 
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000C0&
         Caption         =   "Proses"
         Height          =   735
         Left            =   1560
         TabIndex        =   1
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form_back"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const SYNCHRONIZE       As Long = &H100000
Private Const INFINITE          As Long = &HFFFF
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
    Dim nama As String
    nama = "dbinvent-" & txttgl.Text & ".sql"

    cmd = Chr(34) & "D:\Xampp\MySQL\bin\mysqldump" & Chr(34) & " -uroot  --routines --comments invent_db > c:\Backup\" & nama & ""
    Call execCommand(cmd)

    MsgBox "Backup Database berhasil", vbInformation
End Sub
Private Sub Form_Load()
txttgl.Text = Format(Date, "dd-mm-yy")
End Sub
