VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form_utama 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Barang"
   ClientHeight    =   10695
   ClientLeft      =   60
   ClientTop       =   765
   ClientWidth     =   16950
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_utama.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form_utama.frx":57E2
   ScaleHeight     =   10695
   ScaleWidth      =   16950
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   240
      Top             =   3600
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16935
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   10320
         TabIndex        =   3
         Top             =   1800
         Width           =   7095
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Selamat Datang..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   2235
         End
         Begin VB.Label l_nama 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   2520
            TabIndex        =   4
            Top             =   480
            Width           =   1620
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   4080
         X2              =   15960
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Jl. Siliwangi km 32, Ciutara Parungkuda Sukabumi- Telp 02666722956"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   4320
         TabIndex        =   6
         Top             =   1080
         Width           =   11655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CV. SAKINAH MANDIRI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   840
         Left            =   5760
         TabIndex        =   2
         Top             =   240
         Width           =   7905
      End
      Begin VB.Image Image2 
         Height          =   2460
         Left            =   720
         Picture         =   "Form_utama.frx":1B1C5
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2460
      End
   End
   Begin Crystal.CrystalReport CrUnit 
      Left            =   840
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crDist 
      Left            =   240
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   10200
      Width           =   16950
      _ExtentX        =   29898
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3529
            MinWidth        =   3529
            Text            =   "User"
            TextSave        =   "User"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "6:53 PM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2/13/2014"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport crUser 
      Left            =   1440
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CRTERIMAAL 
      Left            =   120
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crKirimAll 
      Left            =   0
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu dMaster 
      Caption         =   "Data Master"
      Begin VB.Menu tbBarang 
         Caption         =   "Data Barang"
      End
      Begin VB.Menu tbmerk 
         Caption         =   "Data Merk Barang"
      End
      Begin VB.Menu tbjenis 
         Caption         =   "Data Jenis Barang"
      End
      Begin VB.Menu tbdis 
         Caption         =   "Data Distributor"
      End
      Begin VB.Menu tbunit 
         Caption         =   "Data Unit"
      End
   End
   Begin VB.Menu tbstok 
      Caption         =   "Data Stok"
   End
   Begin VB.Menu tbtrans 
      Caption         =   "Transaksi"
      Begin VB.Menu tbterima 
         Caption         =   "Terima Barang"
      End
      Begin VB.Menu tbkirim 
         Caption         =   "Kirim Barang"
      End
   End
   Begin VB.Menu tbret 
      Caption         =   "Retur"
      Begin VB.Menu tbrets 
         Caption         =   "Retur Supplier"
      End
      Begin VB.Menu tbretu 
         Caption         =   "Retur Unit"
      End
   End
   Begin VB.Menu tblap 
      Caption         =   "Laporan"
      Begin VB.Menu lapbar 
         Caption         =   "Laporan Data Barang"
      End
      Begin VB.Menu lapdis 
         Caption         =   "Laporan Data Distributor"
      End
      Begin VB.Menu lapunit 
         Caption         =   "Laporan Data Unit"
      End
      Begin VB.Menu lapuse 
         Caption         =   "Laporan Data User"
      End
      Begin VB.Menu lapterima 
         Caption         =   "Laporan Penerimaan Barang"
         Begin VB.Menu ctk_terima 
            Caption         =   "Cetak Bukti Terima Barang"
         End
         Begin VB.Menu lap_datasep 
            Caption         =   "Laporan Seluruh Data Penerimaan"
         End
         Begin VB.Menu lap_terima_tgl 
            Caption         =   "Laporan Penerimaan PerTanggal"
         End
         Begin VB.Menu lap_penDis 
            Caption         =   "Laporan Penerimaan PerDistributor"
         End
      End
      Begin VB.Menu lapkirim 
         Caption         =   "Laporan Pengiriman Barang"
         Begin VB.Menu ctk_bkti_kirim 
            Caption         =   "Cetak Bukti Pengiriman Barang"
         End
         Begin VB.Menu lap_kiriman 
            Caption         =   "Laporan Seluruh Data Pengiriman"
         End
         Begin VB.Menu lap_kirimtgl 
            Caption         =   "Laporan Pengiriman Per Tanggal"
         End
         Begin VB.Menu rkk 
            Caption         =   "Rekap Pengiriman Per Periode"
         End
      End
      Begin VB.Menu TBReturAl 
         Caption         =   "Laporan Retur Distributor"
      End
      Begin VB.Menu returun 
         Caption         =   "Laporan Retur Unit"
      End
   End
   Begin VB.Menu tbset 
      Caption         =   "Setting"
      Begin VB.Menu tbuser 
         Caption         =   "Tambah User"
      End
      Begin VB.Menu tbubah 
         Caption         =   "Ubah Password"
      End
      Begin VB.Menu tbback 
         Caption         =   "Backup"
      End
   End
   Begin VB.Menu tblog 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "Form_utama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ctes_Click()
Form1.Show 1
End Sub

Private Sub ctk_bkti_kirim_Click()
    ctk_kirim.Show 1
End Sub

Private Sub ctk_terima_Click()
    ctk_BuktiTerima.Show 1
End Sub

Private Sub lap_datasep_Click()
        With Me.CRTERIMAAL
                    .ReportFileName = App.Path & "\Report\TerimaBrg.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    .Action = 1
    End With
End Sub

Private Sub lap_kiriman_Click()
    With Me.crKirimAll
                    .ReportFileName = App.Path & "\Report\kirimunit.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    .Action = 1
    End With
End Sub

Private Sub lap_kirimtgl_Click()
    ctk_kir_un.Show 1
End Sub

Private Sub lap_penDis_Click()
    ctk_terima_dis.Show 1
End Sub

Private Sub lap_terima_tgl_Click()
    Lap_Terima.Show 1
End Sub

Private Sub lapbar_Click()
    Lap_Barang.Show 1
End Sub

Private Sub lapdis_Click()

     Dim SQL As String
        SQL = ""
        SQL = "SELECT * FROM tbdistributor order by kddistributor"
            Set rs_DIS = koneksi.Execute(SQL)
            If rs_DIS.BOF Then
                MsgBox "DATA DISTRIBUTOR TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                "Informasi"
            Else
                With Me.crDist
                    .ReportFileName = App.Path & "\Report\tbDist.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    
                    .Action = 1
                End With
            End If

End Sub



Private Sub lapunit_Click()
  
     Dim SQL As String
        SQL = ""
        SQL = "SELECT * FROM tbunit"
            Set rs_unit = koneksi.Execute(SQL)
            If rs_unit.BOF Then
                MsgBox "DATA UNIT TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                "Informasi"
            Else
                With Me.CrUnit
                    .ReportFileName = App.Path & "\Report\tbunit.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    .Action = 1
                End With
            End If

End Sub

Private Sub lapuse_Click()
     Dim SQL As String
        SQL = ""
        SQL = "SELECT * FROM tbuser"
            Set rs_user = koneksi.Execute(SQL)
            If rs_user.BOF Then
                MsgBox "DATA USER TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                "Informasi"
            Else
                With Me.crUser
                    .ReportFileName = App.Path & "\Report\tbuSER.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    .Action = 1
                End With
            End If
End Sub

Private Sub returun_Click()
    Lap_returUn.Show 1
End Sub

Private Sub rkk_Click()
    Lap_Kirim.Show 1
End Sub

Private Sub tbback_Click()
    Form_back.Show 1
End Sub

Private Sub tbBarang_Click()
    Form_Barang.Show 1
End Sub

Private Sub tbdis_Click()
    Form_distributor.Show 1
End Sub

Private Sub tbjenis_Click()
    Form_MasterJenis.Show 1
End Sub

Private Sub tbkirim_Click()
    Form_LihatKKirim.Show 1
End Sub

Private Sub tblog_Click()
    Pesan = MsgBox("Anda Akan Keluar dari Aplikasi ini ?", vbYesNo + vbQuestion, "Konfirmasi Sistem")
  If Pesan = 6 Then
     Unload Me
     Form_Login.Show 1
  End If

End Sub

Private Sub tbmerk_Click()
    Form_MasterMerk.Show 1
End Sub

Private Sub tbres_Click()
    Form_rest.Show 1
End Sub

Private Sub tbrets_Click()
    Form_LihatReturDis.Show 1
End Sub

Private Sub tbretu_Click()
    Form_LapReturUnit.Show 1
End Sub

Private Sub TBReturAl_Click()
    Lap_ReturSup.Show 1
End Sub

Private Sub tbstok_Click()
    Form_stok.Show 1
End Sub

Private Sub tbterima_Click()
    Form_Lihat_Terima.Show 1
End Sub

Private Sub TbUbah_Click()
    Form_ubahPas.Show 1
End Sub

Private Sub tbunit_Click()
    Form_unit.Show 1
End Sub

Private Sub tbuser_Click()
    Form_user.Show 1
End Sub


Private Sub Timer1_Timer()
If Label3.Visible = True Then
      Label3.Visible = False
   Else
      Label3.Visible = True
   End If
 If l_nama.Visible = True Then
      l_nama.Visible = False
   Else
      l_nama.Visible = True
   End If
End Sub
