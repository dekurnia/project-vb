VERSION 5.00
Begin VB.Form Form_ubahPas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Data Barang"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6855
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
   Icon            =   "Form_ubahPas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUBAH 
      Caption         =   "&UBAH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      Picture         =   "Form_ubahPas.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdKELUAR 
      Caption         =   "&KELUAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      Picture         =   "Form_ubahPas.frx":5E17
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Halaman Ubah Password"
      ForeColor       =   &H00000000&
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtpas3 
         Height          =   390
         Left            =   2520
         TabIndex        =   13
         Top             =   2520
         Width           =   3615
      End
      Begin VB.TextBox txtnama 
         Height          =   390
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtuser 
         Height          =   390
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtpas 
         Height          =   390
         Left            =   2520
         TabIndex        =   2
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox txtpas2 
         Height          =   390
         Left            =   2520
         TabIndex        =   1
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2280
         TabIndex        =   16
         Top             =   2520
         Width           =   60
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2280
         TabIndex        =   15
         Top             =   2040
         Width           =   60
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ulangi Password"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   360
         TabIndex        =   14
         Top             =   2520
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Lengkap"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Width           =   1605
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2280
         TabIndex        =   11
         Top             =   600
         Width           =   60
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2280
         TabIndex        =   9
         Top             =   1080
         Width           =   60
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Password Lama"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   360
         TabIndex        =   8
         Top             =   1560
         Width           =   1710
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2280
         TabIndex        =   7
         Top             =   1560
         Width           =   60
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Password Baru"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   360
         TabIndex        =   6
         Top             =   2040
         Width           =   1605
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         Caption         =   ":"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   9000
         TabIndex        =   5
         Top             =   1080
         Width           =   60
      End
   End
End
Attribute VB_Name = "Form_ubahPas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdKELUAR_Click()
    Unload Me
End Sub

Private Sub cmdUBAH_Click()
    If txtpas.Text = "" Then
        MsgBox "PASSWORD LAMA TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
    ElseIf txtpas2.Text = "" Then
        MsgBox "PASSWORD TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
    ElseIf txtpas3.Text = "" Then
        MsgBox "KONFIRMASI PASSWORD TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
    Else
        SQL = "Select password from tbuser where password=MD5('" & txtpas.Text & "')"
       Set recordset = koneksi.Execute(SQL, , adCmdText)
       If Not recordset.EOF Then
            SQL = "UPDATE tbuser SET password=md5('" & txtpas2 & "') WHERE iduser='" & txtuser.Text & "'"
            koneksi.Execute SQL, , adCmdText
            MsgBox "PASSWORD BERHASIL DIUBAH" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        Else
        MsgBox "PASSWORD LAMA SALAH" + Chr(13) + "NOTE:", 64, "Konfirmasi"
            txtpas.SetFocus
    End If
End If
End Sub

Private Sub Form_Load()
    txtnama.Text = Form_utama.l_nama.Caption
    txtuser.Text = Form_utama.StatusBar1.Panels(1).Text
End Sub

Private Sub txtpas_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If txtpas.Text = "" Then
            MsgBox "PASSWORD LAMA TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
        Else
        txtpas2.SetFocus
        End If
    End If
End Sub



Private Sub txtpas2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If txtpas2.Text = "" Then
            MsgBox "PASSWORD TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
        Else
        txtpas3.SetFocus
        End If
    End If
End Sub



Private Sub txtpas3_KeyPress(KeyAscii As Integer)

KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If txtpas3.Text = "" Then
            MsgBox "PASSWORD 2 TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
        Else
        If txtpas3.Text = txtpas2.Text Then
            cmdUBAH.SetFocus
        Else
            MsgBox "PASSWORD1 TIDAK SAMA DENGAN PASSWORD 2" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
            txtpas3.SetFocus
        End If
        End If
    End If
End Sub
