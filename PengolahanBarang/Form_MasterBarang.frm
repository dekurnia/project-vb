VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_MasterBarang 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sistem Informasi  Pengolahan Barang"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_MasterBarang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame pop1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   2640
      TabIndex        =   29
      Top             =   5400
      Visible         =   0   'False
      Width           =   3855
      Begin MSComctlLib.ListView lvnm1 
         Height          =   3735
         Left            =   0
         TabIndex        =   30
         Top             =   120
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   6588
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nama Barang"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   30
         Left            =   2400
         TabIndex        =   31
         Top             =   840
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   53
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   360
      TabIndex        =   18
      Top             =   5760
      Width           =   8775
      Begin VB.CommandButton cmdlihat 
         Caption         =   "&LIHAT"
         Height          =   855
         Left            =   6240
         Picture         =   "Form_MasterBarang.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   360
         Width           =   1215
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
         Left            =   3840
         Picture         =   "Form_MasterBarang.frx":78A4
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Left            =   240
         Picture         =   "Form_MasterBarang.frx":7F0B
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
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
         Left            =   1440
         Picture         =   "Form_MasterBarang.frx":852D
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdHAPUS 
         Caption         =   "&HAPUS"
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
         Left            =   2640
         Picture         =   "Form_MasterBarang.frx":8B4F
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
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
         Left            =   5040
         Picture         =   "Form_MasterBarang.frx":91C1
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   7560
         Picture         =   "Form_MasterBarang.frx":97F6
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Halaman Master Barang"
      ForeColor       =   &H00000000&
      Height          =   5535
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   8775
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   32
         Top             =   4200
         Width           =   8535
         Begin VB.TextBox txtcari 
            Height          =   390
            Left            =   2160
            TabIndex        =   33
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H000000C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cari Barang"
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   480
            TabIndex        =   34
            Top             =   480
            Width           =   1275
         End
      End
      Begin VB.ComboBox cmbjenis 
         Height          =   390
         Left            =   2760
         TabIndex        =   27
         Top             =   2280
         Width           =   2775
      End
      Begin VB.ComboBox cmbmerk 
         Height          =   390
         Left            =   2760
         TabIndex        =   26
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox cmbsatuan 
         Height          =   390
         Left            =   2760
         TabIndex        =   25
         Text            =   "PCS"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtkode 
         Height          =   390
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtnama 
         Height          =   390
         Left            =   2760
         TabIndex        =   4
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox txtket 
         Height          =   1335
         Left            =   2760
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   2760
         Width           =   3855
      End
      Begin VB.TextBox txtmerk 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtjenis 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   480
         TabIndex        =   17
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   480
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   480
         TabIndex        =   15
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Merk"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   480
         TabIndex        =   14
         Top             =   1800
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   480
         TabIndex        =   13
         Top             =   2280
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2040
         TabIndex        =   12
         Top             =   360
         Width           =   60
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2040
         TabIndex        =   11
         Top             =   840
         Width           =   60
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2040
         TabIndex        =   10
         Top             =   1320
         Width           =   60
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2040
         TabIndex        =   9
         Top             =   1800
         Width           =   60
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2040
         TabIndex        =   8
         Top             =   2280
         Width           =   60
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2040
         TabIndex        =   7
         Top             =   2880
         Width           =   60
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   480
         TabIndex        =   6
         Top             =   2880
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form_MasterBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbjenis_Change()
   If cmbjenis.Text <> "" Then
      Query = "select * from tbjenis where nama='" & cmbjenis.Text & "'"
      Set rsTBJENIS = koneksi.Execute(Query)
      If Not rsTBJENIS.EOF Then
         txtjenis.Text = rsTBJENIS("idjenis")
      End If
   End If
   txtket.SetFocus
End Sub

Private Sub cmbjenis_Click()
 If cmbjenis.Text <> "" Then
      Query = "select * from tbjenis where nama='" & cmbjenis.Text & "'"
      Set rsTBJENIS = koneksi.Execute(Query)
      If Not rsTBJENIS.EOF Then
         txtjenis.Text = rsTBJENIS("idjenis")
      End If
   End If
   txtket.SetFocus
End Sub

Private Sub cmbmerk_Change()
If cmbmerk.Text <> "" Then
      Query = "select * from tbmerk where nama='" & cmbmerk.Text & "'"
      Set rsTBMERK = koneksi.Execute(Query)
      If Not rsTBMERK.EOF Then
         txtmerk.Text = rsTBMERK("idmerk")
      End If
   End If
   cmbjenis.SetFocus
End Sub

Private Sub cmbmerk_Click()
   If cmbmerk.Text <> "" Then
      Query = "select * from tbmerk where nama='" & cmbmerk.Text & "'"
      Set rsTBMERK = koneksi.Execute(Query)
      If Not rsTBMERK.EOF Then
         txtmerk.Text = rsTBMERK("idmerk")
      End If
   End If
   cmbjenis.SetFocus
End Sub

Private Sub cmbsatuan_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Call BlokKarakter(KeyAscii)
    If KeyAscii = 13 Then
        cmbmerk.SetFocus
    End If
End Sub

Private Sub cmdBATAL_Click()
    tdkAktif
    bersih
    
    cmdTAMBAH.Enabled = True
    cmdSIMPAN.Enabled = False
    cmdBATAL.Enabled = False
End Sub



Private Sub cmdHAPUS_Click()
Query = "Call CekBarang('" & txtkode.Text & "')"
    Set rs_BARANG = koneksi.Execute(Query, , adCmdText)
    If Not rs_BARANG.EOF Then
         MsgBox "DATA BARANG TIDAK DAPAT DIHAPUS KARENA DIPAKAI DI TABEL LAIN!", _
         vbInformation, vbOKOnly, "Informasi"
         Exit Sub
    Else
        Query = "CALL HapusBarang('" & txtkode.Text & "')"
        Pesan = MsgBox("Bener Mau Dihapus !" _
            , vbQuestion + vbYesNo, "Konfirmasi")
        If Pesan = vbYes Then
            Set recordset = koneksi.Execute(Query, , adCmdText)
            txtcari.Text = ""
        End If
    End If
End Sub

Private Sub cmdKELUAR_Click()
Unload Me
End Sub

Private Sub cmdlihat_Click()
    Unload Me
    Form_Barang.Show 1
End Sub

Private Sub cmdSIMPAN_Click()
If txtkode.Text = "" Then
    MsgBox "KODE BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    txtkode.SetFocus
 ElseIf txtnama.Text = "" Then
    MsgBox "NAMA BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    txtnama.SetFocus
 ElseIf cmbsatuan.Text = "" Then
     MsgBox "SATUAN BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    cmbsatuan.SetFocus
 ElseIf cmbmerk.Text = "" Then
    MsgBox "MERK BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    cmbmerk.SetFocus
 ElseIf cmbjenis.Text = "" Then
    MsgBox "JENIS BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    cmbjenis.SetFocus
 ElseIf txtket.Text = "" Then
    MsgBox "KETERANGAN BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    txtket.SetFocus
 Else
    If Cek = True Then
       Query = "call TambahBarang('" & txtkode & "','" & txtnama & "','" & cmbsatuan & "','" & txtmerk & "','" & txtjenis & "','" & txtket & "','" & Form_utama.StatusBar1.Panels(1).Text & "',now(),'0','0','0')"
       koneksi.Execute Query, , adCmdText
       MsgBox "DATA BARANG BERHASIL DISIMPAN" + Chr(13) + "NOTE:", 64, "Konfirmasi"
       tdkAktif
       bersih
       Call Form_Activate
    Else
       Query = "call EditBarang('" & txtkode & "','" & txtnama & "','" & cmbsatuan & "','" & txtjenis & "','" & txtmerk & "','" & txtket & "','" & Form_utama.StatusBar1.Panels(1).Text & "',now())"
       koneksi.Execute Query, , adCmdText
       MsgBox "DATA BARANG BERHASIL DIUBAH" + Chr(13) + "NOTE:", 64, "Konfirmasi"
       tdkAktif
       bersih
       Call Form_Activate
    End If
End If
End Sub

Private Sub cmdTAMBAH_Click()
    Cek = True
    Aktifkan
    bersih
    KodeOto
    txtnama.SetFocus
    
    cmdTAMBAH.Enabled = False
    cmdSIMPAN.Enabled = True
    cmdBATAL.Enabled = True
End Sub

Private Sub cmdUBAH_Click()
    Cek = False
    Aktifkan
    cmdSIMPAN.Enabled = True
    txtnama.SetFocus
End Sub

Private Sub Form_Activate()
    pop1.Visible = False
    
         '---Aktifkan Table Merk untuk Combo merk
     cmbmerk.Clear
    Set rsTBMERK = New ADODB.recordset
    rsTBMERK.Open "select nama from tbmerk", koneksi, adOpenDynamic, adLockOptimistic
    Do Until rsTBMERK.EOF
       cmbmerk.AddItem rsTBMERK("nama")
       rsTBMERK.MoveNext
    Loop
    
   '---Aktifkan Table Jenis untuk Combo jenis
   cmbjenis.Clear
    Set rsTBJENIS = New ADODB.recordset
    rsTBJENIS.Open "select nama from tbjenis", koneksi, adOpenDynamic, adLockOptimistic
    Do Until rsTBJENIS.EOF
       cmbjenis.AddItem rsTBJENIS("nama")
       rsTBJENIS.MoveNext
    Loop
    '------------------------
End Sub

Private Sub Form_Click()
pop1.Visible = False
End Sub

Private Sub Form_Load()
   
   Call tdkAktif
   Call bersih
'   frame_br.Visible = False
 '  fr_jenis.Visible = False
End Sub
Private Sub tdkAktif()
    txtnama.Locked = True
    cmbsatuan.Locked = True
    cmbmerk.Locked = True
    cmbjenis.Locked = True
    txtket.Locked = True
    
   cmdHAPUS.Enabled = False
   cmdUBAH.Enabled = False
   cmdSIMPAN.Enabled = False
   cmdBATAL.Enabled = False
   cmdTAMBAH.Enabled = True
   cmdKELUAR.Enabled = True
   cmdlihat.Enabled = True
End Sub
Private Sub bersih()
    txtkode.Text = ""
    txtnama.Text = ""
    cmbsatuan.Text = "PCS"
    cmbmerk.Text = ""
    cmbjenis.Text = ""
    txtket.Text = ""
    txtmerk.Text = ""
    txtjenis.Text = ""
   ' fr_jenis.Visible = False
    'frame_br.Visible = False
End Sub
Public Sub RecTerakhir()
On Error Resume Next
    Query = "select max(kdBarang) from tbbarang"
    Set recordset = koneksi.Execute(Query, , adCmdText)
        If Not recordset.EOF Then
           Me.txtkode = recordset.Fields(0)
        End If
        
End Sub

Sub KodeOto()
    Dim txtkode, KODEMERK As String
    
    RecTerakhir
        If Not Me.txtkode.Text = "" Then
           txtkode = Me.txtkode.Text
           KODEMERK = Val(Right(txtkode, 6) + 1)
            If KODEMERK >= 0 And KODEMERK <= 9 Then
                Me.txtkode.Text = "B-" + "00000" & Trim(Str(KODEMERK))
            ElseIf KODEMERK >= 10 And KODEMERK <= 99 Then
                Me.txtkode.Text = "B-" + "0000" & Trim(Str(KODEMERK))
            ElseIf KODEMERK >= 100 And KODEMERK <= 999 Then
                Me.txtkode.Text = "B-" + "000" & Trim(Str(KODEMERK))
            ElseIf KODEMERK >= 1000 And KODEMERK <= 9999 Then
                Me.txtkode.Text = "B-" + "00" & Trim(Str(KODEMERK))
            ElseIf KODEMERK >= 10000 And KODEMERK <= 99999 Then
                Me.txtkode.Text = "B-" + "0" & Trim(Str(KODEMERK))
            End If
        Else
           Me.txtkode.Text = "B-" + "000001"
        End If
End Sub
Private Sub Aktifkan()
    txtnama.Locked = False
    cmbsatuan.Locked = False
    cmbmerk.Locked = False
    cmbjenis.Locked = False
    txtket.Locked = False
    cmdBATAL.Enabled = True
End Sub

Private Sub satuan()
cmbsatuan.AddItem "KEPING"
cmbsatuan.AddItem "LEMBAR"
cmbsatuan.AddItem "BUAH"
cmbsatuan.AddItem "PCS"
cmbsatuan.AddItem "UNIT"
End Sub


Private Sub namaBarang()
   Query = "call BarangNama('%" & txtcari.Text & "%')"
        Set rs_BARANG = koneksi.Execute(Query, , adCmdText)
        
        If rs_BARANG.EOF Then
            lvnm1.ListItems.Clear
        Else
          rs_BARANG.MoveFirst
                        lvnm1.ListItems.Clear
                        Do While Not rs_BARANG.EOF
                            Set Item = lvnm1.ListItems.Add(, , rs_BARANG.Fields("namaBarang"))
                            rs_BARANG.MoveNext
                        Loop
                        
                        
        End If
End Sub



Private Sub Frame3_Click()
pop1.Visible = False
End Sub

Private Sub lvnm1_DblClick()
        If lvnm1.SelectedItem <> "" Then
                txtcari.Text = lvnm1.SelectedItem
                Query = "call BarangNama('%" & txtcari.Text & "%')"
                Set rs_BARANG = koneksi.Execute(Query, , adCmdText)
                If rs_BARANG.EOF Then
                    MsgBox "DATA TIDAK ADA" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
                Else
                    txtkode.Text = nvl(rs_BARANG.Fields("kdBarang"), "0")
                    txtnama.Text = nvl(rs_BARANG.Fields("namaBarang"), "0")
                    txtcari.Text = nvl(rs_BARANG.Fields("namaBarang"), "0")
                    cmbsatuan.Text = nvl(rs_BARANG.Fields("satuan"), "0")
                    cmbmerk.Text = nvl(rs_BARANG.Fields("merk"), "0")
                    cmbjenis.Text = nvl(rs_BARANG.Fields("jenis"), "0")
                    txtmerk.Text = nvl(rs_BARANG.Fields("idmerk"), "0")
                    txtjenis.Text = nvl(rs_BARANG.Fields("idjenis"), "0")
                    txtket.Text = nvl(rs_BARANG.Fields("keterangan"), "")
                    pop1.Visible = False
                    cmdKELUAR.Enabled = False
                    cmdlihat.Enabled = False
                    cmdSIMPAN.Enabled = False
                    cmdTAMBAH.Enabled = False
                    cmdUBAH.Enabled = True
                    cmdHAPUS.Enabled = True
                End If
    End If
End Sub

Private Sub txtcari_Change()
 pop1.Visible = True
    txtcari.SelStart = Len(txtcari.Text)
    namaBarang
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer)
Call BlokKarakter(KeyAscii)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtket_Change()
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Call BlokKarakter(KeyAscii)
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
Call BlokKarakter(KeyAscii)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Sub
Private Sub txtnama_KeyPress(KeyAscii As Integer)
Call BlokKarakter(KeyAscii)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If txtnama.Text = "" Then
            MsgBox "NAMA BARANG TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
        Else
            cmbsatuan.SetFocus
        End If
    End If
End Sub
Sub Tampilkan()
    On Error Resume Next
    txtkode.Text = recordset.Fields("KDBARANG")
    txtnama.Text = recordset.Fields("NAMABARANG")
    cmbsatuan.Text = recordset.Fields("SATUAN")
    txtmerk.Text = recordset.Fields("IDMERK")
    txtjenis.Text = recordset.Fields("IDJENIS")
    txtket.Text = recordset.Fields("keterangan")
    cmbjenis.Text = recordset.Fields("jenis")
    cmbmerk.Text = recordset.Fields("merk")
End Sub

Public Function nvl(isi, kondisi)
    If IsNull(isi) = True Then
        nvl = kondisi
    Else
        nvl = isi
    End If
End Function

