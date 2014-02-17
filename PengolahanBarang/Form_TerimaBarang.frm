VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form_TerimaBarang 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Stok Barang"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   14550
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
   Icon            =   "Form_TerimaBarang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9945
   ScaleWidth      =   14550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame pop1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   10440
      TabIndex        =   43
      Top             =   1080
      Visible         =   0   'False
      Width           =   3255
      Begin MSComctlLib.ListView lvnm1 
         Height          =   3735
         Left            =   0
         TabIndex        =   47
         Top             =   120
         Width           =   3135
         _ExtentX        =   5530
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
         TabIndex        =   46
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
   Begin VB.Frame Pop 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   2880
      TabIndex        =   39
      Top             =   2160
      Visible         =   0   'False
      Width           =   3375
      Begin MSComctlLib.ListView LvNm 
         Height          =   2655
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   4683
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
            Text            =   "Nama Distributor"
            Object.Width           =   7056
         EndProperty
      End
   End
   Begin VB.TextBox txtKodeBar 
      Height          =   390
      Left            =   840
      TabIndex        =   37
      Top             =   9480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtKodeDist 
      Height          =   390
      Left            =   120
      TabIndex        =   36
      Top             =   9480
      Visible         =   0   'False
      Width           =   615
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
      Picture         =   "Form_TerimaBarang.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   9000
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
      Picture         =   "Form_TerimaBarang.frx":5E77
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   9000
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
      Picture         =   "Form_TerimaBarang.frx":6499
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   9000
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
      Left            =   11760
      Picture         =   "Form_TerimaBarang.frx":6ABB
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9000
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGTERIMA 
      Height          =   2055
      Left            =   360
      TabIndex        =   26
      Top             =   6840
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
   Begin VB.Frame Frame_detail 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detail Barang"
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   360
      TabIndex        =   9
      Top             =   3480
      Width           =   13935
      Begin VB.TextBox txtMar 
         Height          =   390
         Left            =   6000
         TabIndex        =   25
         ToolTipText     =   "Enter Untuk Merubah Harga"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtHargaF 
         Alignment       =   1  'Right Justify
         Height          =   390
         Left            =   2640
         TabIndex        =   24
         ToolTipText     =   "Enter Untuk Mendapatkan Hasil"
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox TxtJmlFix 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtJmlDasar 
         Alignment       =   1  'Right Justify
         Height          =   390
         Left            =   10080
         TabIndex        =   18
         ToolTipText     =   "Enter Untuk Mendapatkan Harga Dasar"
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtQty 
         Height          =   390
         Left            =   2640
         TabIndex        =   15
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox txtHargaDasar 
         Alignment       =   1  'Right Justify
         Height          =   390
         Left            =   2640
         TabIndex        =   12
         ToolTipText     =   "Enter Untuk Mendapatkan Harga Fixed"
         Top             =   840
         Width           =   3135
      End
      Begin VB.Image tbbatal 
         Height          =   480
         Left            =   13080
         Picture         =   "Form_TerimaBarang.frx":7122
         ToolTipText     =   "Cancel"
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label txtlebih 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Rp""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   2
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   62
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Image TbUbah 
         Height          =   480
         Left            =   11880
         Picture         =   "Form_TerimaBarang.frx":7D66
         Stretch         =   -1  'True
         ToolTipText     =   "Ubah"
         Top             =   2520
         Width           =   480
      End
      Begin VB.Image TbHapus 
         Height          =   480
         Left            =   12480
         Picture         =   "Form_TerimaBarang.frx":882A
         ToolTipText     =   "Hapus"
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label Label7 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6600
         TabIndex        =   41
         Top             =   840
         Width           =   255
      End
      Begin VB.Image tbTambah 
         Height          =   495
         Left            =   11280
         Picture         =   "Form_TerimaBarang.frx":946E
         Stretch         =   -1  'True
         ToolTipText     =   "Tambah"
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   23
         Top             =   1320
         Width           =   60
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Fixed"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         Width           =   1290
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   9600
         TabIndex        =   20
         Top             =   1680
         Width           =   60
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga Fixed"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7200
         TabIndex        =   19
         Top             =   1680
         Width           =   2085
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   9600
         TabIndex        =   17
         Top             =   1200
         Width           =   60
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Jml Harga Dasar"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7680
         TabIndex        =   16
         Top             =   1200
         Width           =   1770
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   14
         Top             =   360
         Width           =   60
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Terima"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   11
         Top             =   840
         Width           =   60
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Dasar"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   1350
      End
   End
   Begin VB.Frame frame_iden 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Identifikasi"
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      Begin VB.TextBox txttglfaktur 
         Height          =   390
         Left            =   2640
         TabIndex        =   44
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txttglbukti 
         Height          =   390
         Left            =   2640
         TabIndex        =   42
         Top             =   480
         Width           =   3375
      End
      Begin VB.CommandButton AddDist 
         Height          =   375
         Left            =   6120
         Picture         =   "Form_TerimaBarang.frx":99F6
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtFaktur 
         Height          =   390
         Left            =   2640
         TabIndex        =   33
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox txtDist 
         Height          =   390
         Left            =   2640
         TabIndex        =   8
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox txtkodeBukti 
         Height          =   390
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "yyyy-mm-dd"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   6120
         TabIndex        =   45
         Top             =   2400
         Width           =   1230
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2280
         Picture         =   "Form_TerimaBarang.frx":9DEF
         Stretch         =   -1  'True
         ToolTipText     =   "Find"
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   35
         Top             =   2400
         Width           =   60
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Faktur Reff."
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   34
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   32
         Top             =   1920
         Width           =   60
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "No Faktur Reff."
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   31
         Top             =   1920
         Width           =   1560
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   7
         Top             =   1440
         Width           =   60
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Distributor"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1770
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   4
         Top             =   960
         Width           =   60
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti Terima"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1680
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   2
         Top             =   480
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detail Barang"
      Height          =   3135
      Left            =   7920
      TabIndex        =   48
      Top             =   240
      Width           =   6495
      Begin VB.CommandButton AddBarang 
         Height          =   375
         Left            =   5760
         Picture         =   "Form_TerimaBarang.frx":A0F9
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtBarang 
         Height          =   390
         Left            =   2520
         TabIndex        =   52
         ToolTipText     =   "Masukkan Kode Barang"
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtStok 
         Height          =   390
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   1920
         Width           =   3135
      End
      Begin VB.TextBox txtmerk 
         Height          =   390
         Left            =   2520
         TabIndex        =   50
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtnamaB 
         Height          =   390
         Left            =   2520
         TabIndex        =   49
         ToolTipText     =   "Masukkan Kode Barang"
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   120
         TabIndex        =   60
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2040
         TabIndex        =   59
         Top             =   480
         Width           =   60
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Stok Gudang"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   120
         TabIndex        =   58
         Top             =   1920
         Width           =   1365
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2040
         TabIndex        =   57
         Top             =   1920
         Width           =   60
      End
      Begin VB.Image Image11 
         Height          =   480
         Left            =   2160
         Picture         =   "Form_TerimaBarang.frx":A4F2
         Stretch         =   -1  'True
         ToolTipText     =   "Find"
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2040
         TabIndex        =   56
         Top             =   1440
         Width           =   60
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Merk Barang"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   120
         TabIndex        =   55
         Top             =   1440
         Width           =   1350
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2040
         TabIndex        =   53
         Top             =   960
         Width           =   60
      End
   End
End
Attribute VB_Name = "Form_TerimaBarang"
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
    Form_MasterBarang.Show 1
End Sub

Private Sub AddDist_Click()
    Form_distributor.Show 1
End Sub

Private Sub cmdBATAL_Click()
    SQL = "delete from tbtmpterima where kdkirim='" & txtkodeBukti.Text & "'"
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

Sub tampilgrid()
    Dim BARIS As Integer
    
    FGTERIMA.Clear
    Call AktifGridTerima
     
         
    FGTERIMA.Rows = 2
    BARIS = 0
     
     
   Set rs_TERIMA = New ADODB.recordset
   Query = "Select tbtmpterima.*,tbbarang.namaBarang from tbtmpterima,tbbarang where tbbarang.KdBarang=tbtmpterima.kdBarang and tbtmpterima.kdKirim='" & txtkodeBukti & "'"
   Set rs_TERIMA = koneksi.Execute(Query, , adCmdText)
   
     If rs_TERIMA.EOF Then
     Else
         With rs_TERIMA
            .MoveFirst
         Do While Not .EOF
            BARIS = BARIS + 1
            FGTERIMA.Rows = BARIS + 1
            FGTERIMA.TextMatrix(BARIS, 0) = BARIS
            FGTERIMA.TextMatrix(BARIS, 1) = nvl(.Fields("kdBarang"), "0")
            FGTERIMA.TextMatrix(BARIS, 2) = nvl(.Fields("namaBarang"), "0")
            FGTERIMA.TextMatrix(BARIS, 3) = nvl(.Fields("jumlah"), "0")
            FGTERIMA.TextMatrix(BARIS, 4) = nvl(Format(.Fields("hargaDasar"), "###,###,##0"), "0")
            FGTERIMA.TextMatrix(BARIS, 5) = nvl(Format(.Fields("totalDasar"), "###,###,##0"), "0")
            FGTERIMA.TextMatrix(BARIS, 6) = nvl(Format(.Fields("hargaFixed"), "###,###,##0"), "0")
            FGTERIMA.TextMatrix(BARIS, 7) = nvl(.Fields("persen"), "0")
         .MoveNext
         Loop
         End With
     End If
End Sub

Private Sub cmdSIMPAN_Click()
Dim kurang As Integer
    If txtDist.Text = "" Then
    MsgBox "NAMA DISTRIBUTOR BELUM DIPILIH ", _
        vbOKOnly + vbCritical, "Konfirmasi"
        txtDist.SetFocus
    ElseIf txtFaktur.Text = "" Then
    MsgBox "FAKTUR TIDAK BOLEH KOSONG ", _
        vbOKOnly + vbCritical, "Konfirmasi"
        txtFaktur.SetFocus
    ElseIf BARIS = 1 Then
    MsgBox "Data penerimaan masih kosong", _
        vbOKOnly + vbCritical, "Konfirmasi"
    Else
        Query = "CALL TambahTerimaSem('" & txtkodeBukti & "','" & txttglbukti & "','" & txtFaktur & "','" & txttglfaktur & "','" & txtKodeDist & "','" & Form_utama.StatusBar1.Panels(1).Text & "',now())"
        koneksi.Execute Query, , adCmdText
        
    MsgBox "DATA BERHASIL DISIMPAN DITABEL SEMENTARA, KONFIRM AGAR MENAMBAH STOK" + Chr(13) + "NOTE:", 64, "Konfirmasi"
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
    cmdTAMBAH.Enabled = False
    BARIS = 1
    kosong
End Sub

Private Sub Command1_Click()

End Sub

Private Sub FGTERIMA_DblClick()
    If BARIS = 1 Then
    MsgBox "Data Penerimaan Masih Kosong!", _
        vbOKOnly + vbCritical, "Konfirmasi"
    Else
    
    txtBarang.Enabled = False
    Image1.Enabled = False
    Dim barisGrid As String
    barisGrid = FGTERIMA.Row
    
    If FGTERIMA.Rows <> 1 Then
        txtKodeBar.Text = _
            FGTERIMA.TextMatrix(barisGrid, 1)
        txtBarang.Text = _
            FGTERIMA.TextMatrix(barisGrid, 2)
        txtnamaB.Text = _
            FGTERIMA.TextMatrix(barisGrid, 1)
        txtQty.Text = _
            FGTERIMA.TextMatrix(barisGrid, 3)
        txtHargaDasar.Text = _
            FGTERIMA.TextMatrix(barisGrid, 4)
        txtMar.Text = _
            FGTERIMA.TextMatrix(barisGrid, 7)
        txtHargaF.Text = _
            FGTERIMA.TextMatrix(barisGrid, 6)
        
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
    txtlebih.Caption = Val(Int(txtHargaF.Text)) - Val(Int(txtHargaDasar.Text))
    txtlebih.Caption = Format(txtlebih.Caption, "###,###,##0")
    TxtJmlFix.Text = Val(Int(txtHargaF.Text)) * Val(Int(txtQty.Text))
    TxtJmlFix.Text = Format(TxtJmlFix.Text, "###,###,##0")
    
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
    txttglfaktur.Text = Format(Date, "yyyy-mm-dd")
    Pop.Visible = False
    pop1.Visible = False
    BARIS = 0
End Sub
Private Sub Aktifkan()
    frame_iden.Enabled = True
    Frame_detail.Enabled = True
    FGTERIMA.Enabled = True
    bersih
End Sub
Private Sub tdkAktif()
    frame_iden.Enabled = False
    Frame_detail.Enabled = False
    
    tbTambah.Enabled = False
    
    cmdSIMPAN.Enabled = False
    cmdBATAL.Enabled = False
    cmdKELUAR.Enabled = True
    FGTERIMA.Enabled = False
End Sub

Private Sub bersih()
    txtkodeBukti.Text = ""
    txtDist.Text = ""
    txtFaktur.Text = ""
    
     
    txtBarang.Text = ""
    txtmerk.Text = ""
    txtnamaB.Text = ""
    txtHargaDasar.Text = "0"
    txtQty.Text = "0"
    txtHargaF.Text = "0"
    txtMar.Text = "0"
    txtStok.Text = "0"
    txtJmlDasar.Text = "0"
    TxtJmlFix.Text = "0"
    txtlebih.Caption = ""
    
    
    txtKodeDist.Text = ""
    txtKodeBar.Text = ""
End Sub

Public Sub RecTerakhir()
Dim Query As String
On Error Resume Next
    Query = "select max(kdKirim) from tbtmpTERIMA1"
    Set recordset = koneksi.Execute(Query, , adCmdText)
        If Not recordset.EOF Then
           Me.txtkodeBukti.Text = recordset.Fields(0)
        End If
        
End Sub

Sub KodeOto()
Dim txtNOBM As String
Dim NOBM

RecTerakhir
    If Not Me.txtkodeBukti.Text = "" Then
       txtNOBM = Me.txtkodeBukti.Text
       NOBM = Val(Left(txtNOBM, 4) + 1)
        If NOBM >= 0 And NOBM <= 9 Then
               Me.txtkodeBukti.Text = "000" & Trim(Str(NOBM)) + "/" + "BM" + "/" + Mid(txttglbukti.Text, 6, 2) + "/" + Right(Date, 2)
        ElseIf NOBM >= 10 And NOBM <= 99 Then
               Me.txtkodeBukti.Text = "00" & Trim(Str(NOBM)) + "/" + "BM" + "/" + Mid(txttglbukti.Text, 6, 2) + "/" + Right(Date, 2)
        ElseIf NOBM >= 100 And NOBM <= 999 Then
               Me.txtkodeBukti.Text = "0" & Trim(Str(NOBM)) + "/" + "BM" + "/" + Mid(txttglbukti.Text, 6, 2) + "/" + Right(Date, 2)
        ElseIf NOBM >= 1000 And NOBM <= 9999 Then
               Me.txtkodeBukti.Text = Trim(Str(NOBM)) + "/" + "BM" + "/" + Mid(txttglbukti.Text, 6, 2) + "/" + Right(Date, 2)
        End If
    Else
       Me.txtkodeBukti.Text = "0001" + "/" + "BM" + "/" + Mid(txttglbukti.Text, 6, 2) + "/" + Right(Date, 2)
    End If
End Sub

Private Sub Frame_detail_Click()
pop1.Visible = False
End Sub



Private Sub frame_iden_Click()
Pop.Visible = False
End Sub


Private Sub Frame1_Click()
    pop1.Visible = False
End Sub

Private Sub Image1_Click()
        Pop.Visible = True
        Query = "select * from tbdistributor ORDER BY kdDistributor"
        Set recordset = koneksi.Execute(Query, , adCmdText)
        If recordset.EOF Then
            LvNm.ListItems.Clear
        Else
          recordset.MoveFirst
                        LvNm.ListItems.Clear
                        Do While Not recordset.EOF
                            Set Item = LvNm.ListItems.Add(, , recordset.Fields("namaDistributor"))
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
                            Set Item = lvnm1.ListItems.Add(, , recordset.Fields("namaBarang"))
                            recordset.MoveNext
                        Loop
                        
        End If
End Sub

Private Sub LvNm_Click()
    If LvNm.SelectedItem <> "" Then
                txtDist.Text = LvNm.SelectedItem
                Query = "select * from tbdistributor WHERE namaDistributor = '" & txtDist.Text & "' ORDER BY kdDistributor"
                Set rs_DIS = koneksi.Execute(Query, , adCmdText)
                If rs_DIS.EOF Then
                    MsgBox "DATA TIDAK ADA" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
                Else
                    txtDist.Text = nvl(rs_DIS.Fields("namaDistributor"), "0")
                    txtKodeDist.Text = nvl(rs_DIS.Fields("kdDistributor"), "0")
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
                    txtnamaB.Text = nvl(rs_BARANG.Fields("kdBarang"), "0")
                    txtKodeBar.Text = nvl(rs_BARANG.Fields("kdBarang"), "0")
                    txtStok.Text = nvl(rs_BARANG.Fields("stokAkhir"), "0")
                    txtmerk.Text = nvl(rs_BARANG.Fields("merk"), "0")
                    pop1.Visible = False
                    txtQty.SetFocus
                End If
    End If
End Sub



Private Sub tbbatal_Click()
    Call kosong
    TbHapus.Enabled = False
    TbUbah.Enabled = False
    tbTambah.Enabled = True
End Sub

Private Sub TbHapus_Click()
    Query = "CALL HapusTmpTer('" & txtkodeBukti.Text & "','" & txtKodeBar.Text & "')"
        Pesan = MsgBox("Bener Mau Dihapus !" _
            , vbQuestion + vbYesNo, "Konfirmasi")
        If Pesan = vbYes Then
            Set recordset = koneksi.Execute(Query, , adCmdText)
            Call kosong
            Me.FGTERIMA.Refresh
            Call tampilgrid
            TbUbah.Enabled = False
                TbHapus.Enabled = False
                tbbatal.Enabled = False
                tbTambah.Enabled = True
                txtBarang.Enabled = True
                Image1.Enabled = True
                BARIS = BARIS - 1
        End If
End Sub

Private Sub tbTambah_Click()
    If txtBarang.Text = "" Then
        MsgBox "NAMA BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        txtBarang.SetFocus
    ElseIf txtQty.Text = "0" Then
        MsgBox "KUANTITI BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        txtQty.SetFocus
    ElseIf txtHargaDasar.Text = "0" Then
        MsgBox "HARGA DASAR BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        txtHargaDasar.SetFocus
    ElseIf txtJmlDasar.Text = "0" Then
        MsgBox "JUMLAH HARGA DASAR TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    ElseIf txtHargaF.Text = "0" Then
        MsgBox "HARGA FIX TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    ElseIf TxtJmlFix.Text = "0" Then
        MsgBox "JUMLAH HARGA FIX TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    ElseIf txtMar.Text = "0" Then
        MsgBox "PERSEN TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        txtMar.SetFocus
    Else
        Pesan = MsgBox("YAKIN MAU DIINPUT ??" _
            , vbQuestion + vbYesNo, "Konfirmasi")
        If Pesan = vbYes Then
        
            SQL = "select * from tbtmpterima where kdKirim='" & txtkodeBukti & "' and kdBarang='" & txtKodeBar & "'"
            Set rs_BARANG = koneksi.Execute(SQL, , adCmdText)
            If Not rs_BARANG.EOF Then
                 MsgBox "DATA BARANG SUDAH ADA, TIDAK BISA DIINPUT LAGI!", _
                 vbInformation + vbOKOnly, "Informasi"
                 Exit Sub
            Else
                Query = "call TambahTmpKir('" & txtkodeBukti & "','" & txtKodeBar & "','" & txtQty.Text & "','" & Int(txtHargaDasar) & "','" & Int(txtJmlDasar) & "','" & Int(txtHargaF) & "','" & Int(txtMar) & "')"
                koneksi.Execute Query, , adCmdText
                Call kosong
                Me.FGTERIMA.Refresh
                Call tampilgrid
                BARIS = BARIS + 1
                txtHargaDasar.SetFocus
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
        MsgBox "HARGA DASAR BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        txtHargaDasar.SetFocus
    ElseIf txtJmlDasar.Text = "" Then
        MsgBox "JUMLAH HARGA DASAR TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    ElseIf txtHargaF.Text = "" Then
        MsgBox "HARGA FIX TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    ElseIf TxtJmlFix.Text = "" Then
        MsgBox "JUMLAH HARGA FIX TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    ElseIf txtMar.Text = "" Then
        MsgBox "PERSEN TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        txtMar.SetFocus
    Else
        Query = "call EditTmpTer('" & txtkodeBukti & "','" & txtKodeBar & "','" & txtQty.Text & "','" & Int(txtHargaDasar) & "','" & Int(txtJmlDasar) & "','" & Int(txtHargaF) & "','" & Int(txtMar) & "')"
        koneksi.Execute Query, , adCmdText
                Call kosong
                Me.FGTERIMA.Refresh
                Call tampilgrid
                TbUbah.Enabled = False
                TbHapus.Enabled = False
                tbTambah.Enabled = True
                txtBarang.Enabled = True
                Image1.Enabled = True
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
    txtDist.SelStart = Len(txtDist.Text)
    namaDistributor
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
Private Sub namaDistributor()
    Query = "select * from tbdistributor WHERE namaDistributor LIKE '%" & txtDist.Text & "%' ORDER BY kdDistributor"
        Set recordset = koneksi.Execute(Query, , adCmdText)
        If recordset.EOF Then
            LvNm.ListItems.Clear
        Else
          recordset.MoveFirst
                        LvNm.ListItems.Clear
                        Do While Not recordset.EOF
                            Set Item = LvNm.ListItems.Add(, , recordset.Fields("namaDistributor"))
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
Private Sub txtFaktur_KeyPress(KeyAscii As Integer)
Call BlokKarakter(KeyAscii)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If txtFaktur.Text = "" Then
            MsgBox "FAKTUR REFERENSI TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
        Else
            txttglfaktur.SetFocus
        End If
    End If
End Sub

Private Sub txtHargaDasar_Change()
    If txtHargaDasar.Text = "" Then
        txtHargaDasar.Text = "0"
    End If
End Sub

Private Sub txtHargaDasar_KeyPress(KeyAscii As Integer)
    Call HanyaNomor(KeyAscii)
    If KeyAscii = 13 Then
        If txtHargaDasar.Text = "" Or txtHargaDasar.Text = "0" Or txtQty.Text = "" Or txtQty.Text = "0" Then
            MsgBox "HARGA DASAR TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
            txtHargaDasar.SetFocus
        Else
       
       txtJmlDasar.Text = Val(Int(txtHargaDasar.Text)) * Val(Int(txtQty.Text))
       txtJmlDasar.Text = Format(txtJmlDasar.Text, "###,###,##0")
    
       txtHargaF.Text = (Val(Int(txtHargaDasar.Text)) * Val(Int(txtMar.Text) / 100))
       txtHargaF.Text = Format(txtHargaF.Text, "###,###,##0")
       
       txtlebih.Caption = (Val(Int(txtHargaDasar.Text)) * Val(Int(txtMar.Text) / 100))
       txtlebih.Caption = Format(txtlebih.Caption, "###,###,##0")
       
       TxtJmlFix.Text = Val(Int(txtHargaF.Text)) * Val(Int(txtQty.Text))
       TxtJmlFix.Text = Format(TxtJmlFix.Text, "###,###,##0")
       txtHargaDasar.Text = Format(txtHargaDasar.Text, "###,###,##0")
       End If
    End If
End Sub

Private Sub txtHargaF_Change()
    txtHargaF.Text = Format(txtHargaF.Text, "###,###,##0")
    txtHargaF.SelStart = Len(txtHargaF.Text)
    If txtHargaF.Text = "" Then
        txtHargaF.Text = 0
    End If
End Sub

Private Sub txtHargaF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMar.Text = (Val(Int(txtHargaDasar.Text)) * Val(Int(txtMar.Text) / 100))
        TxtJmlFix.Text = Val(Int(txtHargaF.Text)) * Val(Int(txtQty.Text))
        TxtJmlFix.Text = Format(TxtJmlFix.Text, "###,###,##0")
    End If
End Sub



Private Sub txtJmlDasar_KeyPress(KeyAscii As Integer)
    Call HanyaNomor(KeyAscii)
    If KeyAscii = 13 Then
        If txtJmlDasar.Text = "" Or txtQty.Text = "" Or txtQty.Text = "0" Or txtJmlDasar.Text = "0" Then
            txtJmlDasar.SetFocus
        Else
        txtHargaDasar.Text = Val(Int(txtJmlDasar.Text)) / Val(Int(txtQty.Text))
        txtHargaDasar.Text = Format(txtHargaDasar.Text, "###,###,##0")
        txtJmlDasar.Text = Format(txtJmlDasar.Text, "###,###,##0")
        
        txtHargaF.Text = (Val(Int(txtHargaDasar.Text)) * Val(Int(txtMar.Text) / 100))
       txtHargaF.Text = Format(txtHargaF.Text, "###,###,##0")
       
       txtlebih.Caption = (Val(Int(txtHargaDasar.Text)) * Val(Int(txtMar.Text) / 100))
       txtlebih.Caption = Format(txtlebih.Caption, "###,###,##0")
       
       
       TxtJmlFix.Text = Val(Int(txtHargaF.Text)) * Val(Int(txtQty.Text))
       TxtJmlFix.Text = Format(TxtJmlFix.Text, "###,###,##0")
        End If
    End If
End Sub



Private Sub txtMar_Change()
    If txtMar.Text = "" Then
        txtMar.Text = "0"
    End If
End Sub

Private Sub txtMar_KeyPress(KeyAscii As Integer)
    Call HanyaNomor(KeyAscii)
    If KeyAscii = 13 Then
        If txtMar.Text = "" Then
        Else
            If txtHargaDasar.Text = "" Or txtHargaDasar.Text = "0" Or txtJmlDasar.Text = "" Or txtJmlDasar.Text = "0" Then
            Else
            txtHargaF.Text = (Val(Int(txtHargaDasar.Text)) * Val(Int(txtMar.Text) / 100))
            txtHargaF.Text = Format(txtHargaF.Text, "###,###,##0")
            
            TxtJmlFix.Text = Val(Int(txtHargaF.Text)) * Val(Int(txtQty.Text))
            TxtJmlFix.Text = Format(TxtJmlFix.Text, "###,###,##0")
            
            txtlebih.Caption = (Val(Int(txtHargaDasar.Text)) * Val(Int(txtMar.Text) / 100))
            txtlebih.Caption = Format(txtlebih.Caption, "###,###,##0")
       
            End If
        End If
    End If
End Sub



Private Sub txtQty_KeyPress(KeyAscii As Integer)
    Call HanyaNomor(KeyAscii)
     If KeyAscii = 13 Then
        If txtQty.Text = "" Or txtQty.Text = "0" Then
            MsgBox "KUANTITY TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
            txtQty.SetFocus
        Else
            If Not txtHargaDasar.Text = "" Or txtHargaDasar.Text = "0" Or txtJmlDasar.Text = "" Or txtJmlDasar.Text = "0" Then
                txtJmlDasar.Text = Val(Int(txtHargaDasar.Text)) * Val(Int(txtQty.Text))
                txtJmlDasar.Text = Format(txtJmlDasar.Text, "###,###,##0")
                
                TxtJmlFix.Text = Val(Int(txtHargaF.Text)) * Val(Int(txtQty.Text))
                TxtJmlFix.Text = Format(TxtJmlFix.Text, "###,###,##0")
                
                txtMar.Text = 25
                txtHargaDasar.SetFocus
            Else
                txtMar.Text = 25
                txtHargaDasar.SetFocus
            End If
        End If
    End If
End Sub
Private Sub namaBarang()
   Query = "call BarangNama('%" & txtBarang.Text & "%')"
        Set rs_BARANG = koneksi.Execute(Query, , adCmdText)
        'lvnm1.ListItems(lvnm1.FindItem(storedvalue).Index).EnsureVisible
        If rs_BARANG.EOF Then
            lvnm1.ListItems.Clear
        Else
          rs_BARANG.MoveFirst
                        lvnm1.ListItems.Clear
                        Do While Not rs_BARANG.EOF
                            Set Item = lvnm1.ListItems.Add(, , rs_BARANG.Fields("namaBarang"))
                            rs_BARANG.MoveNext
                        Loop
                        'lvnm1.SelectedItem.EnsureVisible
                        
        End If
End Sub
Private Sub kosong()
   txtBarang.Text = ""
    txtmerk.Text = ""
    txtnamaB.Text = ""
    txtHargaDasar.Text = "0"
    txtQty.Text = "0"
    txtHargaF.Text = "0"
    txtMar.Text = "0"
    txtStok.Text = "0"
    txtJmlDasar.Text = "0"
    TxtJmlFix.Text = "0"
    txtlebih.Caption = "0"
    
    txtKodeBar.Text = ""
End Sub

Private Sub AktifGridTerima()
    With FGTERIMA
        
        .RowHeightMin = 300
        .Col = 0
        .Row = 0
        .Text = "NO"
        .CellFontBold = True
        .ColWidth(0) = 300
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 1
        .Row = 0
        .Text = "KODE BARANG"
        .CellFontBold = True
        .ColWidth(1) = 1500
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 2
        .Row = 0
        .Text = "NAMA BARANG"
        .CellFontBold = True
        .ColWidth(2) = 3000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 3
        .Row = 0
        .Text = "KUANTITI"
        .CellFontBold = True
        .ColWidth(3) = 1000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 4
        .Row = 0
        .Text = "HARGA"
        .CellFontBold = True
        .ColWidth(4) = 1500
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 5
        .Row = 0
        .Text = "TOTAL HARGA"
        .CellFontBold = True
        .ColWidth(5) = 2000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 6
        .Row = 0
        .Text = "HARGA FIXED"
        .CellFontBold = True
        .ColWidth(6) = 2000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 7
        .Row = 0
        .Text = "TAMBAHAN"
        .CellFontBold = True
        .ColWidth(7) = 1000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
               
    End With
End Sub
Function ListViewScroll(lvnm1 As ListView, ByVal dx As Long, ByVal dy As Long)
    SendMessage lvnm1.hwnd, LVM_SCROLL, dx, ByVal dy
    SendMessage LvNm.hwnd, LVM_SCROLL, dx, ByVal dy
End Function

Private Sub txttglfaktur_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtBarang.SetFocus
End If
End Sub
