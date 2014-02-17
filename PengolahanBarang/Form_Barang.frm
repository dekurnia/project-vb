VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form_Barang 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pegolahan Data Barang"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   14910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Barang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   14910
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGBARANG 
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   11668
      _Version        =   393216
      Cols            =   10
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
      _Band(0).Cols   =   10
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Data Barang"
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14655
      Begin VB.TextBox txtcari 
         Height          =   390
         Left            =   11160
         TabIndex        =   3
         Top             =   480
         Width           =   2775
      End
      Begin VB.ComboBox cmbcari 
         Height          =   390
         Left            =   9240
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   555
         Left            =   13920
         Picture         =   "Form_Barang.frx":57E2
         ToolTipText     =   "Refresh"
         Top             =   360
         Width           =   555
      End
      Begin VB.Image TbRefresh 
         Height          =   480
         Left            =   1920
         Picture         =   "Form_Barang.frx":5E9B
         ToolTipText     =   "Refresh"
         Top             =   480
         Width           =   480
      End
      Begin VB.Image TbHapus 
         Height          =   480
         Left            =   1080
         Picture         =   "Form_Barang.frx":6ADF
         ToolTipText     =   "Hapus"
         Top             =   480
         Width           =   480
      End
      Begin VB.Image tbTambah 
         Height          =   495
         Left            =   240
         Picture         =   "Form_Barang.frx":7723
         Stretch         =   -1  'True
         ToolTipText     =   "Tambah"
         Top             =   480
         Width           =   480
      End
   End
End
Attribute VB_Name = "Form_Barang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub FGBARANG_DblClick()
    Dim barisGrid As String
    barisGrid = FGBARANG.Row
    
    If FGBARANG.Rows <> 1 Then
        txtcari.Text = _
            FGBARANG.TextMatrix(barisGrid, 1)
    Else
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
cmbcari.AddItem "Kode Barang"
cmbcari.AddItem "Nama Barang"
cmbcari.AddItem "Jenis Barang"
cmbcari.AddItem "Merk Barang"

Call TampilGrid
End Sub

Private Sub TbHapus_Click()
If txtcari.Text = "" Then
    MsgBox "KODE BARANG KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
Else
    Query = "Call CekBarang('" & txtcari.Text & "')"
    Set rs_BARANG = koneksi.Execute(Query, , adCmdText)
    If Not rs_BARANG.EOF Then
         MsgBox "DATA BARANG TIDAK DAPAT DIHAPUS KARENA DIPAKAI DI TABEL LAIN!", _
         vbInformation + vbOKOnly, "Informasi"
         Exit Sub
    Else
        Query = "CALL HapusBarang('" & txtcari.Text & "')"
        Pesan = MsgBox("Bener Mau Dihapus !" _
            , vbQuestion + vbYesNo, "Konfirmasi")
        If Pesan = vbYes Then
            Set recordset = koneksi.Execute(Query, , adCmdText)
            txtcari.Text = ""
        End If
    End If
End If
End Sub

Private Sub TbRefresh_Click()
    Me.FGBARANG.Refresh
    Call TampilGrid
    txtcari.Text = ""
    cmbcari.Text = ""
End Sub

Private Sub tbTambah_Click()
    Unload Me
    Form_MasterBarang.Show 1
End Sub
Sub TampilGrid()
    Dim BARIS As Integer
    
    FGBARANG.Clear
    Call AktifGridBarang
     
         
    FGBARANG.Rows = 2
    BARIS = 0
     
     
   Set rs_BARANG = New ADODB.recordset
   If cmbcari.Text = "Nama Barang" Then
        Query = "call BarangNama('%" & txtcari.Text & "%')"
   ElseIf cmbcari.Text = "Kode Barang" Then
        Query = "call BarangKode('%" & txtcari.Text & "%')"
   ElseIf cmbcari.Text = "Jenis Barang" Then
        Query = "call BarangJenis('%" & txtcari.Text & "%')"
   ElseIf cmbcari.Text = "Merk Barang" Then
        Query = "call BarangMerk('%" & txtcari.Text & "%')"
   Else
        Query = "call TampilBarang()"
   End If
   Set rs_BARANG = koneksi.Execute(Query, , adCmdText)
   
     If rs_BARANG.EOF Then
         MsgBox "DATA KOSONG!", _
         vbInformation + vbOKOnly, "Informasi"
         Exit Sub
     Else
         With rs_BARANG
            .MoveFirst
         Do While Not .EOF
            
            BARIS = BARIS + 1
            FGBARANG.Rows = BARIS + 1
            FGBARANG.TextMatrix(BARIS, 0) = BARIS
            FGBARANG.TextMatrix(BARIS, 1) = .Fields("kdbarang")
            FGBARANG.TextMatrix(BARIS, 2) = .Fields("namabarang")
            FGBARANG.TextMatrix(BARIS, 3) = .Fields("MERK")
            FGBARANG.TextMatrix(BARIS, 4) = .Fields("JENIS")
            FGBARANG.TextMatrix(BARIS, 5) = .Fields("keterangan")
            FGBARANG.TextMatrix(BARIS, 6) = .Fields("stokAkhir")
            FGBARANG.TextMatrix(BARIS, 7) = .Fields("Satuan")
            FGBARANG.TextMatrix(BARIS, 8) = .Fields("user_ubah")
            FGBARANG.TextMatrix(BARIS, 9) = .Fields("tgl_ubah")
         .MoveNext
         Loop
         End With
     End If
End Sub
Sub AktifGridBarang()
    With FGBARANG
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
        .Text = "MERK BARANG"
        .CellFontBold = True
        .ColWidth(3) = 2000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 4
        .Row = 0
        .Text = "JENIS BARANG"
        .CellFontBold = True
        .ColWidth(4) = 2000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 5
        .Row = 0
        .Text = "KETERANGAN"
        .CellFontBold = True
        .ColWidth(5) = 3000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 6
        .Row = 0
        .Text = "STOK AKHIR"
        .CellFontBold = True
        .ColWidth(6) = 2000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 7
        .Row = 0
        .Text = "SATUAN"
        .CellFontBold = True
        .ColWidth(7) = 2000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 8
        .Row = 0
        .Text = "USER UBAH"
        .CellFontBold = True
        .ColWidth(8) = 2000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 9
        .Row = 0
        .Text = "TGL UBAH"
        .CellFontBold = True
        .ColWidth(9) = 2000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
    End With
End Sub
Public Function nvl(isi, kondisi)
    If IsNull(isi) = True Then
        nvl = kondisi
    Else
        nvl = isi
    End If
End Function

Private Sub txtcari_Change()
   If cmbcari.Text = "Nama Barang" Then
        Query = "call BarangNama('%" & txtcari.Text & "%')"
   ElseIf cmbcari.Text = "Kode Barang" Then
        Query = "call BarangKode('%" & txtcari.Text & "%')"
   ElseIf cmbcari.Text = "Jenis Barang" Then
        Query = "call BarangJenis('%" & txtcari.Text & "%')"
   ElseIf cmbcari.Text = "Merk Barang" Then
        Query = "call BarangMerk('%" & txtcari.Text & "%')"
   ElseIf cmbcari.Text = "" Then
   Query = "call TampilBarang()"
   End If
     Set recordset = koneksi.Execute(Query, , adCmdText)
        If recordset.EOF Then
            MsgBox "TIDAK MENEMUKAN KODE JENIS! " _
            & " - " & txtcari.Text & " - dalam tabel", _
            vbInformation, "Informasi"
            
            txtcari.Text = ""
            txtcari.SetFocus
        Else
          Call TampilGrid
        End If
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer)
    Call BlokKarakter(KeyAscii)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
