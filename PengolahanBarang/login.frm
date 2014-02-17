VERSION 5.00
Begin VB.Form Form_Login 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Barang"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "login.frx":57E2
   ScaleHeight     =   3900
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMASUK 
      Caption         =   "&MASUK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      Picture         =   "login.frx":79E5
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdBATAL 
      Caption         =   "&BATAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      Picture         =   "login.frx":7F88
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Halaman User"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   5775
      Begin VB.ComboBox cmbLevel 
         Height          =   390
         Left            =   1440
         TabIndex        =   6
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox txtPas 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox txtUser 
         Height          =   495
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Level"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Username"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1080
      End
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   360
      Picture         =   "login.frx":85E6
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "Form_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MaxLogin As Integer
Private Sub cmbLevel_Change()
    cmdMASUK.SetFocus
End Sub
Private Sub cmblevel_Click()
    cmdMASUK.SetFocus
End Sub
Private Sub cmdBATAL_Click()
    End
End Sub
Private Sub cmdMASUK_Click()
Dim level As String
If txtuser.Text = "" Then
MsgBox "USER TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
            txtuser.SetFocus
ElseIf txtpas.Text = "" Then
MsgBox "PASSWORD TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
            txtpas.SetFocus
ElseIf cmblevel.Text = "" Then
MsgBox "LEVEL TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
            cmblevel.SetFocus
Else
    Query = "Select * from tbUSER Where iduser ='" & txtuser & "' AND password=MD5('" & txtpas & "') AND level ='" & cmblevel & "' AND status='Y'"
        Set recordset = koneksi.Execute(Query, , adCmdText)
        If Not recordset.EOF Then
            level = recordset.Fields("level")
            MsgBox "SELAMAT ANDA BERHASIL LOGIN" + Chr(13) + "INFORMASI", 64, ""
            Unload Me
                Form_utama.l_nama.Caption = recordset.Fields("namalengkap")
                Form_utama.StatusBar1.Panels(1).Text = recordset.Fields("iduser")
                Form_utama.StatusBar1.Panels(2).Text = recordset.Fields("level")
            If level = "ADMIN" Then
                Form_utama.dMaster.Enabled = True
                Form_utama.tbstok.Enabled = False
                Form_utama.lapuse.Enabled = False
                Form_utama.tbuser.Enabled = False
                Form_utama.tbback.Enabled = False
                Form_utama.Show 1
            ElseIf level = "MANAGER" Then
                Form_utama.dMaster.Enabled = False
                Form_utama.tbtrans.Enabled = False
                Form_utama.tbret.Enabled = False
                Form_utama.tbuser.Enabled = False
                Form_utama.tbback.Enabled = False
                Form_utama.Show 1
            ElseIf level = "ADMINISTRATOR" Then
                Form_utama.Show 1
            End If
        Else
            If MaxLogin < 3 Then
                MsgBox "USER/PASSWORD MASIH SALAH, SILAHKAN ULANGI LAGI!", _
                    vbCritical + vbOKOnly, "Error"
                txtuser.Text = ""
                txtpas.Text = ""
                cmblevel.Text = ""
                txtuser.SetFocus
                MaxLogin = MaxLogin + 1
            Else
                MsgBox "ANDA BUKAN USER YANG BERHAK!", _
                    vbCritical + vbOKOnly, "Error"
                End
            End If
        End If
End If
End Sub
Private Sub Form_Load()
    Call BukaDatabase
    Call level
End Sub
Private Sub txtpas_KeyPress(KeyAscii As Integer)
Call BlokKarakter(KeyAscii)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If txtpas.Text = "" Then
            MsgBox "PASSWORD TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
        Else
            cmblevel.SetFocus
        End If
    End If
End Sub

Private Sub txtuser_KeyPress(KeyAscii As Integer)
Call BlokKarakter(KeyAscii)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If txtuser.Text = "" Then
            MsgBox "USERNAME TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
        Else
            txtpas.SetFocus
        End If
    End If
End Sub
Private Sub level()
    cmblevel.AddItem "ADMINISTRATOR"
    cmblevel.AddItem "ADMIN"
    cmblevel.AddItem "MANAGER"
End Sub
