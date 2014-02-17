VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form_user 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Data Barang"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   13575
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Form_user.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   13575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   240
      TabIndex        =   16
      Top             =   3240
      Width           =   13095
      Begin VB.CommandButton cmdSIMPAN 
         Caption         =   "&SIMPAN"
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
         Left            =   1320
         Picture         =   "Form_user.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdBATAL 
         Caption         =   "&BATAL"
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
         Left            =   2520
         Picture         =   "Form_user.frx":5E04
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdTAMBAH 
         Caption         =   "&TAMBAH"
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
         Left            =   120
         Picture         =   "Form_user.frx":646B
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSIMPAN1 
         Caption         =   "&SIMPAN"
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
         Left            =   1320
         Picture         =   "Form_user.frx":6A8D
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   -1680
         Width           =   1095
      End
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
         Left            =   3720
         Picture         =   "Form_user.frx":70AF
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
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
         Left            =   4920
         Picture         =   "Form_user.frx":76E4
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtcari 
         Height          =   390
         Left            =   9600
         TabIndex        =   17
         Top             =   600
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   555
         Left            =   12360
         Picture         =   "Form_user.frx":7D79
         ToolTipText     =   "Refresh"
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Masukkan Username"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7200
         TabIndex        =   24
         Top             =   600
         Width           =   2205
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Halaman User"
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   13095
      Begin VB.CheckBox Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   9240
         TabIndex        =   28
         Top             =   2040
         Width           =   615
      End
      Begin VB.ComboBox cmblevel 
         Height          =   390
         Left            =   9240
         TabIndex        =   15
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox txtpas2 
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   9240
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtpas 
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   9240
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtuser 
         Height          =   390
         Left            =   2520
         TabIndex        =   5
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtnama 
         Height          =   390
         Left            =   2520
         TabIndex        =   2
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Status Aktif"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7080
         TabIndex        =   27
         Top             =   2040
         Width           =   1185
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   9000
         TabIndex        =   26
         Top             =   2040
         Width           =   60
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   9000
         TabIndex        =   14
         Top             =   1560
         Width           =   60
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7080
         TabIndex        =   13
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   9000
         TabIndex        =   12
         Top             =   1080
         Width           =   60
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ulangi Password"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7080
         TabIndex        =   10
         Top             =   1080
         Width           =   1770
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   9000
         TabIndex        =   9
         Top             =   600
         Width           =   60
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7080
         TabIndex        =   7
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2280
         TabIndex        =   6
         Top             =   1080
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
         TabIndex        =   4
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2280
         TabIndex        =   3
         Top             =   600
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Lengkap"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1605
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGDIST 
      Height          =   3855
      Left            =   240
      TabIndex        =   25
      Top             =   4680
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   7
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
End
Attribute VB_Name = "Form_user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmblevel_Click()
    cmdSIMPAN.SetFocus
End Sub

Private Sub cmdBATAL_Click()
    Call TampilGrid
    Call tdkAktif
    Call bersih
    txtnama.Locked = False
    txtuser.Locked = False
End Sub



Private Sub cmdKELUAR_Click()
Unload Me
End Sub

Private Sub cmdSIMPAN_Click()
Dim ST As String
If Option1.Value = Checked Then
    ST = "Y"
Else
    ST = "N"
End If

 If txtnama.Text = "" Then
    MsgBox "NAMA TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    txtnama.SetFocus
 ElseIf txtuser.Text = "" Then
    MsgBox "USERNAME TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    txtuser.SetFocus
 ElseIf txtpas.Text = "" Then
     MsgBox "PASSWORD TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    txtpas.SetFocus
 ElseIf cmblevel.Text = "" Then
    MsgBox "PILIH LEVEL" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    cmblevel.SetFocus
 Else
    If Cek = True Then
       SQL = "Select iduser from tbuser where iduser='" & txtuser.Text & "'"
       Set recordset = koneksi.Execute(SQL, , adCmdText)
       If Not recordset.EOF Then
            MsgBox "USERNAME SUDAH ADA, SILAHKAN GUNAKAN USERNAME YANG LAIN" + Chr(13) + "NOTE:", 64, "Konfirmasi"
            txtuser.SetFocus
        Else
            Query = "call TambahuSER('" & txtnama & "','" & txtuser & "',md5('" & txtpas & "'),'" & cmblevel & "','" & Form_utama.StatusBar1.Panels(1).Text & "',now(),'" & ST & "')"
            koneksi.Execute Query, , adCmdText
            MsgBox "DATA USER BERHASIL DISIMPAN" + Chr(13) + "NOTE:", 64, "Konfirmasi"
            Call Form_Activate
            Me.FGDIST.Refresh
        End If
    Else
        
            Query = "call EditUser('" & txtnama & "','" & txtuser & "',md5('" & txtpas & "'),'" & cmblevel & "','" & Form_utama.StatusBar1.Panels(1).Text & "',now(),'" & ST & "')"
            koneksi.Execute Query, , adCmdText
            MsgBox "DATA USER BERHASIL DIUBAH" + Chr(13) + "NOTE:", 64, "Konfirmasi"
            Call Form_Activate
            Me.FGDIST.Refresh
            txtnama.Locked = False
    txtuser.Locked = False
        
    End If
End If
End Sub

Private Sub cmdTAMBAH_Click()
    Cek = True
    Aktifkan
    bersih
    txtnama.SetFocus
    
    cmdTAMBAH.Enabled = False
    cmdSIMPAN.Enabled = True
    cmdUBAH.Enabled = False
    'cmdHAPUS.Enabled = False
    cmdBATAL.Enabled = True
End Sub



Private Sub cmdUBAH_Click()
    Aktifkan
    Cek = False
End Sub

Private Sub FGDIST_DblClick()
'    Aktifkan
    Dim barisGrid As String
    Dim ST As String
    barisGrid = FGDIST.Row
    
    If FGDIST.Rows <> 1 Then
        txtnama.Text = _
            FGDIST.TextMatrix(barisGrid, 1)
        txtuser.Text = _
            FGDIST.TextMatrix(barisGrid, 2)
        cmblevel.Text = _
            FGDIST.TextMatrix(barisGrid, 3)
        ST = _
            FGDIST.TextMatrix(barisGrid, 6)
            
    Else
        Exit Sub
    End If
    If ST = "Y" Then
        Option1.Value = Checked
    Else
        Option1.Value = Unchecked
    End If
    txtnama.Locked = True
    txtuser.Locked = True
    cmdUBAH.Enabled = True
    'cmdHAPUS.Enabled = True
    cmdBATAL.Enabled = True
    cmdKELUAR.Enabled = False
End Sub

Private Sub Form_Activate()
    Call TampilGrid
    Call tdkAktif
    Call bersih
    
    cmblevel.AddItem "ADMINISTRATOR"
    cmblevel.AddItem "ADMIN"
    cmblevel.AddItem "MANAGER"
End Sub
Private Sub tdkAktif()
   Frame1.Enabled = False
   'cmdHAPUS.Enabled = False
   cmdUBAH.Enabled = False
   cmdSIMPAN.Enabled = False
   cmdBATAL.Enabled = False
   cmdTAMBAH.Enabled = True
   cmdKELUAR.Enabled = True
End Sub
Private Sub Aktifkan()
   Frame1.Enabled = True
   'cmdHAPUS.Enabled = True
   cmdUBAH.Enabled = True
   cmdSIMPAN.Enabled = True
   cmdBATAL.Enabled = True
   cmdTAMBAH.Enabled = False
   cmdKELUAR.Enabled = False
End Sub
Private Sub bersih()
    txtnama.Text = ""
    txtuser.Text = ""
    txtpas.Text = ""
    txtpas2.Text = ""
    cmblevel.Text = ""
    Option1.Value = False
End Sub
Sub TampilGrid()
    Dim BARIS As Integer
    
    FGDIST.Clear
    Call AktifGridDis
     
         
    FGDIST.Rows = 2
    BARIS = 0
     
     
   Set rs_DIS = New ADODB.recordset
   Query = "select * from tbuser WHERE IDUSER LIKE '%" & txtcari.Text & "%'"
   Set rs_DIS = koneksi.Execute(Query, , adCmdText)
   
     If rs_DIS.EOF Then
         MsgBox "DATA KOSONG!", _
         vbInformation + vbOKOnly, "Informasi"
         Exit Sub
     Else
         With rs_DIS
            .MoveFirst
         Do While Not .EOF
            BARIS = BARIS + 1
            FGDIST.Rows = BARIS + 1
            FGDIST.TextMatrix(BARIS, 0) = BARIS
            FGDIST.TextMatrix(BARIS, 1) = nvl(.Fields("namalengkap"), "0")
            FGDIST.TextMatrix(BARIS, 2) = nvl(.Fields("IDUSER"), "0")
            FGDIST.TextMatrix(BARIS, 3) = nvl(.Fields("level"), "0")
            FGDIST.TextMatrix(BARIS, 4) = nvl(.Fields("user_ubah"), "0")
            FGDIST.TextMatrix(BARIS, 5) = nvl(.Fields("tgl_ubah"), "0")
            FGDIST.TextMatrix(BARIS, 6) = nvl(.Fields("status"), "0")
         .MoveNext
         Loop
         End With
     End If
End Sub
Sub AktifGridDis()
    With FGDIST
        .RowHeightMin = 300
        .Col = 0
        .Row = 0
        .Text = "NO"
        .CellFontBold = True
        .ColWidth(0) = 400
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .RowHeightMin = 300
        .Col = 1
        .Row = 0
        .Text = "NAMA LENGKAP"
        .CellFontBold = True
        .ColWidth(1) = 2500
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 2
        .Row = 0
        .Text = "USERNAME"
        .CellFontBold = True
        .ColWidth(2) = 2500
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 3
        .Row = 0
        .Text = "LEVEL"
        .CellFontBold = True
        .ColWidth(3) = 2000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        
        
        .Col = 4
        .Row = 0
        .Text = "USER UBAH"
        .CellFontBold = True
        .ColWidth(4) = 2000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 5
        .Row = 0
        .Text = "TGL UBAH"
        .CellFontBold = True
        .ColWidth(5) = 2000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
         .Col = 6
        .Row = 0
        .Text = "STATUS"
        .CellFontBold = True
        .ColWidth(6) = 1000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
    End With
End Sub



Private Sub txtcari_Change()
    Query = "select * from tbuser WHERE IDUSER LIKE '%" & txtcari.Text & "%'"
     Set recordset = koneksi.Execute(Query, , adCmdText)
        If recordset.EOF Then
            MsgBox "TIDAK MENEMUKAN NAMA DISTRIBUTOR! " _
            & " - " & txtcari.Text & " - dalam tabel", _
            vbInformation, "Informasi"
            
            txtcari.Text = ""
            txtcari.SetFocus
        Else
          Call TampilGrid
        End If
End Sub

Private Sub txtCARI_KeyPress(KeyAscii As Integer)
Call BlokKarakter(KeyAscii)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtnama_KeyPress(KeyAscii As Integer)
    Call BlokKarakter(KeyAscii)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If txtnama.Text = "" Then
            MsgBox "NAMA LENGKAP TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
        Else
            txtuser.SetFocus
        End If
    End If
End Sub


Private Sub txtpas_KeyPress(KeyAscii As Integer)
Call BlokKarakter(KeyAscii)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If txtpas.Text = "" Then
            MsgBox "PASSWORD TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
        Else
        txtpas2.SetFocus
        End If
    End If
End Sub

Private Sub txtpas2_KeyPress(KeyAscii As Integer)
Call BlokKarakter(KeyAscii)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If txtpas2.Text = "" Then
            MsgBox "PASSWORD 2 TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
        Else
        If txtpas.Text = txtpas2.Text Then
            cmblevel.SetFocus
        Else
            MsgBox "PASSWORD1 TIDAK SAMA DENGAN PASSWORD 2" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
        End If
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
Public Function nvl(isi, kondisi)
    If IsNull(isi) = True Then
        nvl = kondisi
    Else
        nvl = isi
    End If
End Function
