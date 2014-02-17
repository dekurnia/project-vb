VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form_KirimBarang 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Data Barang"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   14220
   ControlBox      =   0   'False
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
   Icon            =   "Form_KirimBarang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   14220
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Pop 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   2760
      TabIndex        =   43
      Top             =   1800
      Visible         =   0   'False
      Width           =   3375
      Begin MSComctlLib.ListView LvNm 
         Height          =   2415
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   4260
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nama Unit"
            Object.Width           =   7056
         EndProperty
      End
   End
   Begin VB.Frame pop1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   2760
      TabIndex        =   35
      Top             =   2760
      Visible         =   0   'False
      Width           =   3135
      Begin MSComctlLib.ListView lvnm1 
         Height          =   3495
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   6165
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nama Barang"
            Object.Width           =   7056
         EndProperty
      End
   End
   Begin VB.TextBox txtKodeDist 
      Height          =   390
      Left            =   240
      TabIndex        =   34
      Top             =   7560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtKodeBar 
      Height          =   390
      Left            =   960
      TabIndex        =   33
      Top             =   7560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame frame_iden 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Identifikasi"
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   13935
      Begin VB.TextBox txtkodeBukti 
         Height          =   390
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtDist 
         Height          =   390
         Left            =   2640
         TabIndex        =   25
         Top             =   1440
         Width           =   3375
      End
      Begin VB.CommandButton AddDist 
         Height          =   375
         Left            =   6120
         Picture         =   "Form_KirimBarang.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txttglbukti 
         Height          =   390
         Left            =   2640
         TabIndex        =   23
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   32
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   31
         Top             =   480
         Width           =   60
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti Kirim"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   30
         Top             =   960
         Width           =   1500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   29
         Top             =   960
         Width           =   60
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Unit"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   27
         Top             =   1440
         Width           =   60
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2280
         Picture         =   "Form_KirimBarang.frx":5BDB
         Stretch         =   -1  'True
         ToolTipText     =   "Find"
         Top             =   1440
         Width           =   480
      End
   End
   Begin VB.Frame Frame_detail 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detail Barang"
      ForeColor       =   &H00000000&
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   13935
      Begin VB.TextBox txtmerk 
         Height          =   390
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox namaB 
         Height          =   390
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtBarang 
         Height          =   390
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtHargaDasar 
         Height          =   390
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtQty 
         Height          =   390
         Left            =   10080
         TabIndex        =   9
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox txtJmlDasar 
         Height          =   390
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtStok 
         Height          =   390
         Left            =   2640
         TabIndex        =   7
         Top             =   2160
         Width           =   3135
      End
      Begin VB.CommandButton AddBarang 
         Height          =   375
         Left            =   5880
         Picture         =   "Form_KirimBarang.frx":5EE5
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.Image tbbatal 
         Height          =   480
         Left            =   12960
         Picture         =   "Form_KirimBarang.frx":62DE
         ToolTipText     =   "Cancel"
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   42
         Top             =   1200
         Width           =   60
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Merk Barang"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   41
         Top             =   1200
         Width           =   1350
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   39
         Top             =   720
         Width           =   60
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   38
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   20
         Top             =   360
         Width           =   60
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Harga"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   19
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   18
         Top             =   1680
         Width           =   60
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kuantitas"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7680
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   9600
         TabIndex        =   16
         Top             =   960
         Width           =   60
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7200
         TabIndex        =   15
         Top             =   1440
         Width           =   1440
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   9600
         TabIndex        =   14
         Top             =   1440
         Width           =   60
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Stok Gudang"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   1365
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   12
         Top             =   2160
         Width           =   60
      End
      Begin VB.Image tbTambah 
         Height          =   495
         Left            =   11160
         Picture         =   "Form_KirimBarang.frx":6F22
         Stretch         =   -1  'True
         ToolTipText     =   "Tambah"
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image Image11 
         Height          =   480
         Left            =   2280
         Picture         =   "Form_KirimBarang.frx":74AA
         Stretch         =   -1  'True
         ToolTipText     =   "Find"
         Top             =   360
         Width           =   480
      End
      Begin VB.Image TbHapus 
         Height          =   480
         Left            =   12360
         Picture         =   "Form_KirimBarang.frx":77B4
         ToolTipText     =   "Hapus"
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image TbUbah 
         Height          =   480
         Left            =   11760
         Picture         =   "Form_KirimBarang.frx":83F8
         Stretch         =   -1  'True
         ToolTipText     =   "Ubah"
         Top             =   2040
         Width           =   480
      End
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
      Left            =   11760
      Picture         =   "Form_KirimBarang.frx":8EBC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
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
      Left            =   9360
      Picture         =   "Form_KirimBarang.frx":9523
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
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
      Left            =   10560
      Picture         =   "Form_KirimBarang.frx":9B45
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
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
      Left            =   12960
      Picture         =   "Form_KirimBarang.frx":A167
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7080
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGTERIMA 
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
End
Attribute VB_Name = "Form_KirimBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BARIS As Integer
Dim i As Integer

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Const LVM_FIRST = &H1000
Const LVM_SCROLL = (LVM_FIRST + 20)
Private Sub AddBarang_Click()
    Form_Barang.Show 1
End Sub

Private Sub AddDist_Click()
    Form_unit.Show 1
End Sub

Private Sub cmdBATAL_Click()
        SQL = "delete from tbtmpkirimUn where kdkirimun='" & txtkodeBukti.Text & "'"
        Pesan = MsgBox("Bener Neeh Mau Dibatalin !" _
                , vbQuestion + vbYesNo, "Konfirmasi")
        If Pesan = vbYes Then
           Set recordset = koneksi.Execute(SQL, adCmdText)
           Call Form_Load
           Me.FGTERIMA.Refresh
            cmdTAMBAH.Enabled = True
        End If
End Sub

Private Sub cmdKELUAR_Click()
Unload Me
End Sub

Private Sub cmdSIMPAN_Click()
    Dim kurang As Integer
    If txtDist.Text = "" Then
    MsgBox "NAMA UNIT BELUM DIPILIH ", _
        vbOKOnly + vbCritical, "Konfirmasi"
        txtDist.SetFocus
    ElseIf BARIS = 1 Then
    MsgBox "Data pengiriman masih kosong", _
        vbOKOnly + vbCritical, "Konfirmasi"
    Else
        Query = "CALL TambahKirimUnSem('" & txtkodeBukti & "','" & txttglbukti & "','" & txtKodeDist & "','" & Form_utama.StatusBar1.Panels(1).Text & "',now())"
        koneksi.Execute Query, , adCmdText
    MsgBox "DATA BERHASIL DISIMPAN DI TABEL SEMENTARA, SILAHKAN KONFIRM UNTUK MEMPENGARUHI STOK" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    Call Form_Load
    End If
End Sub

Private Sub cmdTAMBAH_Click()
    Cek = True
    Aktifkan
    KodeOto
    tampilgrid
    tbTambah.Enabled = True
    cmdSIMPAN.Enabled = True
    cmdBATAL.Enabled = True
    cmdKELUAR.Enabled = False
    BARIS = 1
    kosong
    tbTambah.Enabled = True
End Sub



Private Sub cmdUBAH_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub FGTERIMA_DblClick()
If BARIS = 1 Then
Else
    Dim barisGrid As String
    txtBarang.Enabled = False
    Image1.Enabled = False
    barisGrid = FGTERIMA.Row
    
    If FGTERIMA.Rows <> 1 Then
        txtKodeBar.Text = _
            FGTERIMA.TextMatrix(barisGrid, 1)
        txtBarang.Text = _
            FGTERIMA.TextMatrix(barisGrid, 2)
        namaB.Text = _
            FGTERIMA.TextMatrix(barisGrid, 1)
        txtQty.Text = _
            FGTERIMA.TextMatrix(barisGrid, 3)
        txtHargaDasar.Text = _
            FGTERIMA.TextMatrix(barisGrid, 4)
        txtJmlDasar.Text = _
            FGTERIMA.TextMatrix(barisGrid, 5)
    Else
        Exit Sub
    End If
    SQL2 = "Select *From tbBARANG where KDbarang ='" & txtKodeBar & "'"
        Set rs_BARANG = koneksi.Execute(SQL2, , adCmdText)
        If Not rs_BARANG.EOF Then
            txtStok.Text = rs_BARANG.Fields("stokAkhir")
        Else
        End If
    txtJmlDasar.Text = Val(Int(txtHargaDasar.Text)) * Val(Int(txtQty.Text))
    txtJmlDasar.Text = Format(txtJmlDasar.Text, "###,###,##0")
    
    TbUbah.Enabled = True
    TbHapus.Enabled = True
    tbTambah.Enabled = False
    tbbatal.Enabled = True
End If
End Sub

Private Sub Form_Click()
    Pop.Visible = False
    pop1.Visible = False
End Sub

Private Sub Form_Load()
    tdkAktif
    bersih
    FGTERIMA.Clear
    txttglbukti.Text = Format(Date, "yyyy-mm-dd")
    Pop.Visible = False
    pop1.Visible = False
    BARIS = 0
End Sub
Private Sub tdkAktif()
    Frame_iden.Enabled = False
    Frame_detail.Enabled = False
    
    tbTambah.Enabled = False
    cmdSIMPAN.Enabled = False
    cmdBATAL.Enabled = False
    cmdKELUAR.Enabled = True
    cmdTAMBAH.Enabled = True
    FGTERIMA.Enabled = False
End Sub
Private Sub Aktifkan()
    Frame_iden.Enabled = True
    Frame_detail.Enabled = True
    
    tbTambah.Enabled = True
    cmdSIMPAN.Enabled = True
    cmdBATAL.Enabled = True
    cmdKELUAR.Enabled = False
    cmdTAMBAH.Enabled = False
    FGTERIMA.Enabled = True
End Sub

Private Sub bersih()
    txtkodeBukti.Text = ""
    txtDist.Text = ""
     
    txtBarang.Text = ""
    txtHargaDasar.Text = "0"
    txtQty.Text = "0"
    txtStok.Text = "0"
    txtJmlDasar.Text = "0"
    namaB.Text = ""
    txtmerk.Text = ""
    
    txtKodeDist.Text = ""
    txtKodeBar.Text = ""
End Sub
Sub tampilgrid()
    Dim BARIS As Integer
    
    FGTERIMA.Clear
    Call AktifGridTerima
     
         
    FGTERIMA.Rows = 2
    BARIS = 0
     
     
   Set rs_TERIMA = New ADODB.recordset
   Query = "Select tbtmpKirimUn.*,tbbarang.namaBarang from tbtmpKirimUn,tbbarang where tbbarang.KdBarang=tbtmpKirimUn.kdBarang and tbtmpKirimUn.kdKirimUn='" & txtkodeBukti & "'"
   Set rs_TERIMA = koneksi.Execute(Query, , adCmdText)
   
     If rs_TERIMA.EOF Then
     Else
         With rs_TERIMA
            .MoveFirst
         Do While Not .EOF
            BARIS = BARIS + 1
            FGTERIMA.Rows = BARIS + 1
            FGTERIMA.TextMatrix(BARIS, 0) = BARIS
            FGTERIMA.TextMatrix(BARIS, 1) = .Fields("kdBarang")
            FGTERIMA.TextMatrix(BARIS, 2) = .Fields("namaBarang")
            FGTERIMA.TextMatrix(BARIS, 3) = .Fields("jml")
            FGTERIMA.TextMatrix(BARIS, 4) = Format(.Fields("harga"), "###,###,##0")
            FGTERIMA.TextMatrix(BARIS, 5) = Format(.Fields("total"), "###,###,##0")
         .MoveNext
         Loop
         End With
     End If
End Sub
Private Sub kosong()
    txtBarang.Text = ""
    txtHargaDasar.Text = 0
    txtQty.Text = 0
    txtStok.Text = 0
    txtJmlDasar.Text = 0
    namaB.Text = ""
    txtmerk.Text = ""
    
    
    txtKodeBar.Text = ""
End Sub

Private Sub AktifGridTerima()
    With FGTERIMA
        
        .RowHeightMin = 300
        .Col = 0
        .Row = 0
        .Text = "NO"
        .CellFontBold = True
        .ColWidth(0) = 0
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 1
        .Row = 0
        .Text = "KODE BARANG"
        .CellFontBold = True
        .ColWidth(1) = 2000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 2
        .Row = 0
        .Text = "NAMA BARANG"
        .CellFontBold = True
        .ColWidth(2) = 4000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 3
        .Row = 0
        .Text = "KUANTITI"
        .CellFontBold = True
        .ColWidth(3) = 2000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 4
        .Row = 0
        .Text = "HARGA DASAR"
        .CellFontBold = True
        .ColWidth(4) = 2500
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 5
        .Row = 0
        .Text = "TOTAL HARGA"
        .CellFontBold = True
        .ColWidth(5) = 3000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        
               
    End With
End Sub

Private Sub Frame_detail_Click()
    pop1.Visible = False
End Sub

Private Sub frame_iden_Click()
    Pop.Visible = False
End Sub

Public Sub RecTerakhir()
Dim Query As String
On Error Resume Next
    Query = "select max(kdKirimUn) from tbtmpkirimun1"
    Set recordset = koneksi.Execute(Query, , adCmdText)
        If Not recordset.EOF Then
           Me.txtkodeBukti.Text = recordset.Fields(0)
        End If
        
End Sub

Sub KodeOto()
Dim txtNOBM As String
Dim NOBM

Call RecTerakhir
    If Not Me.txtkodeBukti.Text = "" Then
       txtNOBM = Me.txtkodeBukti.Text
       NOBM = Val(Left(txtNOBM, 4) + 1)
        If NOBM >= 0 And NOBM <= 9 Then
               Me.txtkodeBukti.Text = "000" & Trim(Str(NOBM)) + "/" + "KU" + "/" + Mid(txttglbukti.Text, 6, 2) + "/" + Right(Date, 2)
        ElseIf NOBM >= 10 And NOBM <= 99 Then
               Me.txtkodeBukti.Text = "00" & Trim(Str(NOBM)) + "/" + "KU" + "/" + Mid(txttglbukti.Text, 6, 2) + "/" + Right(Date, 2)
        ElseIf NOBM >= 100 And NOBM <= 999 Then
               Me.txtkodeBukti.Text = "0" & Trim(Str(NOBM)) + "/" + "KU" + "/" + Mid(txttglbukti.Text, 6, 2) + "/" + Right(Date, 2)
        ElseIf NOBM >= 1000 And NOBM <= 9999 Then
               Me.txtkodeBukti.Text = Trim(Str(NOBM)) + "/" + "KU" + "/" + Mid(txttglbukti.Text, 6, 2) + "/" + Right(Date, 2)
        End If
    Else
       Me.txtkodeBukti.Text = "0001" + "/" + "KU" + "/" + Mid(txttglbukti.Text, 6, 2) + "/" + Right(Date, 2)
    End If
End Sub

Private Sub Image1_Click()
        Pop.Visible = True
        Query = "select * from tbUnit ORDER BY kdUnit"
        Set recordset = koneksi.Execute(Query, , adCmdText)
        If recordset.EOF Then
            LvNm.ListItems.Clear
        Else
          recordset.MoveFirst
                        LvNm.ListItems.Clear
                        Do While Not recordset.EOF
                            Set Item = LvNm.ListItems.Add(, , recordset.Fields("namaUnit"))
                            recordset.MoveNext
                        Loop
                        
        End If
End Sub

Private Sub Image11_Click()
            pop1.Visible = True
        Query = "select * from tbbarang ORDER BY kdBarang"
        Set recordset = koneksi.Execute(Query, , adCmdText)
        If recordset.EOF Then
            lvnm1.ListItems.Clear
        Else
          recordset.MoveFirst
                        lvnm1.ListItems.Clear
                        Do While Not recordset.EOF
                            Set Item = lvnm1.ListItems.Add(, , recordset.Fields("kdBarang"))
                            recordset.MoveNext
                        Loop
                        
        End If
End Sub

Private Sub LvNm_Click()
        If LvNm.SelectedItem <> "" Then
                txtDist.Text = LvNm.SelectedItem
                Query = "select * from tbunit WHERE namaunit = '" & txtDist.Text & "' ORDER BY kdunit"
                Set rs_DIS = koneksi.Execute(Query, , adCmdText)
                If rs_DIS.EOF Then
                    MsgBox "DATA TIDAK ADA" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
                Else
                    txtDist.Text = nvl(rs_DIS.Fields("namaunit"), "0")
                    txtKodeDist.Text = nvl(rs_DIS.Fields("kdunit"), "0")
                    Pop.Visible = False
                End If
    End If
End Sub



Private Sub lvnm1_Click()
     If lvnm1.SelectedItem <> "" Then
                txtBarang.Text = lvnm1.SelectedItem
                Query = "call BarangNama('%" & txtBarang.Text & "%')"
                Set rs_BARANG = koneksi.Execute(Query, , adCmdText)
                If rs_BARANG.EOF Then
                    MsgBox "DATA TIDAK ADA" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
                Else
                    txtBarang.Text = nvl(rs_BARANG.Fields("namaBarang"), "0")
                    namaB.Text = nvl(rs_BARANG.Fields("kdBarang"), "0")
                    txtmerk.Text = nvl(rs_BARANG.Fields("merk"), "0")
                    txtKodeBar.Text = nvl(rs_BARANG.Fields("kdBarang"), "0")
                    txtStok.Text = nvl(rs_BARANG.Fields("stokAkhir"), "0")
                    txtHargaDasar.Text = Format(nvl(rs_BARANG.Fields("HargaFixed"), "0"), "###,###,##0")
                    pop1.Visible = False
                    txtQty.SetFocus
                End If
              
    End If
End Sub

Private Sub tbbatal_Click()
Call kosong
TbUbah.Enabled = False
TbHapus.Enabled = False
tbTambah.Enabled = True
End Sub

Private Sub TbHapus_Click()
        Query = "CALL HapusTmpKirUn('" & txtkodeBukti.Text & "','" & txtKodeBar.Text & "')"
        Pesan = MsgBox("Bener Mau Dihapus !" _
            , vbQuestion + vbYesNo, "Konfirmasi")
        If Pesan = vbYes Then
            Set recordset = koneksi.Execute(Query, , adCmdText)
            Call kosong
            Me.FGTERIMA.Refresh
            Call tampilgrid
            TbUbah.Enabled = False
                TbHapus.Enabled = False
                tbTambah.Enabled = True
                txtBarang.Enabled = True
               BARIS = BARIS - 1
                Image1.Enabled = True
        End If
End Sub

Private Sub tbTambah_Click()
    If txtBarang.Text = "" Then
        MsgBox "NAMA BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        txtBarang.SetFocus
    ElseIf txtQty.Text = "" Then
        MsgBox "KUANTITI BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        txtQty.SetFocus
    ElseIf txtHargaDasar.Text = "" Then
        MsgBox "HARGA  BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        txtHargaDasar.SetFocus
    ElseIf txtJmlDasar.Text = "" Then
        MsgBox "TOTAL HARGA TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    ElseIf Val(txtQty.Text) > Val(txtStok.Text) Then
        MsgBox "JUMLAH KIRIM TIDAK BOLEH LEBIH BESAR DARI STOK BARANG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    Else
        Pesan = MsgBox("YAKIN MAU DIINPUT ??" _
            , vbQuestion + vbYesNo, "Konfirmasi")
        If Pesan = vbYes Then
            SQL = "select * from tbtmpkirimun where kdKirimUn='" & txtkodeBukti & "' and kdBarang='" & txtKodeBar & "'"
            Set rs_BARANG = koneksi.Execute(SQL, , adCmdText)
            If Not rs_BARANG.EOF Then
                 MsgBox "DATA BARANG SUDAH ADA, TIDAK BISA DIINPUT LAGI!", _
                 vbInformation + vbOKOnly, "Informasi"
                 Exit Sub
            Else
                Query = "call TambahTmpKirUn('" & txtkodeBukti & "','" & txtKodeBar & "','" & txtQty.Text & "','" & Int(txtHargaDasar) & "','" & Int(txtJmlDasar) & "')"
                koneksi.Execute Query, , adCmdText
                Call kosong
                Me.FGTERIMA.Refresh
                Call tampilgrid
                txtHargaDasar.SetFocus
                BARIS = BARIS + 1
            End If
        End If
    End If
End Sub

Private Sub TbUbah_Click()
        If txtBarang.Text = "" Then
        MsgBox "NAMA BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        txtBarang.SetFocus
    ElseIf txtQty.Text = "" Then
        MsgBox "KUANTITI BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        txtQty.SetFocus
    ElseIf txtHargaDasar.Text = "" Then
        MsgBox "HARGA  BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        txtHargaDasar.SetFocus
    ElseIf txtJmlDasar.Text = "" Then
        MsgBox "TOTAL HARGA TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
   Else
        Pesan = MsgBox("YAKIN MAU DIINPUT ??" _
            , vbQuestion + vbYesNo, "Konfirmasi")
        If Pesan = vbYes Then
                sqleditki = "UPDATE tbtmpkirimun SET jml='" & txtQty.Text & "',harga='" & Int(txtHargaDasar) & "',total='" & Int(txtJmlDasar) & "' WHERE" _
                            & " kdkirimun='" & txtkodeBukti & "' and kdbarang='" & txtKodeBar & "'"
    
                koneksi.Execute sqleditki, , adCmdText
                Call kosong
                Me.FGTERIMA.Refresh
                Call tampilgrid
                txtHargaDasar.SetFocus
                txtBarang.Enabled = True
                Image1.Enabled = True
        End If
    End If
End Sub

Private Sub txtBarang_Change()
    pop1.Visible = True
    txtBarang.SelStart = Len(txtBarang.Text)
    namaBarang
End Sub

Private Sub txtBarang_KeyPress(KeyAscii As Integer)
Call BlokKarakter(KeyAscii)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDist_Change()
        Pop.Visible = True
        pop1.Visible = False
        txtDist.SelStart = Len(txtDist.Text)
        namaUnit
End Sub
Private Sub namaUnit()
    Query = "select * from tbunit WHERE namaunit LIKE '%" & txtDist.Text & "%' ORDER BY kdunit"
        Set recordset = koneksi.Execute(Query, , adCmdText)
        If recordset.EOF Then
            LvNm.ListItems.Clear
        Else
          recordset.MoveFirst
                        LvNm.ListItems.Clear
                        Do While Not recordset.EOF
                            Set Item = LvNm.ListItems.Add(, , recordset.Fields("namaunit"))
                            recordset.MoveNext
                        Loop
                        
        End If
End Sub
Public Function nvl(isi, kondisi)
    If IsNull(isi) = True Then
        nvl = kondisi
    Else
        nvl = isi
    End If
End Function
Private Sub namaBarang()
   Query = "call BarangNama('%" & txtBarang.Text & "%')"
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
Private Sub txtDist_KeyPress(KeyAscii As Integer)
Call BlokKarakter(KeyAscii)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDist_KeyUp(KeyCode As Integer, Shift As Integer)
    Call BlokKarakter(KeyAscii)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyCode = vbKeyF1 Then
       Pop.Visible = True
    End If
End Sub

Private Sub txtQty_Change()
Call HanyaNomor(KeyAscii)
 txtQty.SelStart = Len(txtQty.Text)
 If txtQty.Text = "" Then
        txtQty.Text = 0
    End If
 txtJmlDasar.Text = Val(Int(txtHargaDasar.Text)) * Val(Int(txtQty.Text))
 txtJmlDasar.Text = Format(txtJmlDasar.Text, "###,###,##0")
 
End Sub

Function ListViewScroll(lvnm1 As ListView, ByVal dx As Long, ByVal dy As Long)
    SendMessage lvnm1.hwnd, LVM_SCROLL, dx, ByVal dy
    SendMessage LvNm.hwnd, LVM_SCROLL, dx, ByVal dy
End Function
