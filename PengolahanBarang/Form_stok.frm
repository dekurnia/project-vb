VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form_stok 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Data Barang"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   14925
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
   Icon            =   "Form_stok.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   14925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Data Barang"
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   14655
      Begin Crystal.CrystalReport CRSTOK 
         Left            =   360
         Top             =   720
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox txtcari 
         Height          =   390
         Left            =   10800
         TabIndex        =   7
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtstok 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Image TbCetak 
         Height          =   510
         Left            =   8640
         Picture         =   "Form_stok.frx":57E2
         ToolTipText     =   "Klik Untuk Mencetak"
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Stok Akhir"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   5760
         TabIndex        =   6
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   840
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Masukan Kode Barang Yang Anda Cari"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   6480
         TabIndex        =   2
         Top             =   600
         Width           =   4110
      End
      Begin VB.Image TbRefresh 
         Height          =   480
         Left            =   14040
         Picture         =   "Form_stok.frx":78A4
         ToolTipText     =   "Refresh"
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   555
         Left            =   13440
         Picture         =   "Form_stok.frx":84E8
         ToolTipText     =   "Klik Untuk Mencari"
         Top             =   360
         Width           =   555
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGBARANG 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   11456
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
End
Attribute VB_Name = "Form_stok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
     '---Aktifkan Table Merk untuk Combo merk
     txtCARI.Clear
    Set rsTBMERK = New ADODB.recordset
    rsTBMERK.Open "select kdbarang from tbbarang order by kdbarang", koneksi, adOpenDynamic, adLockOptimistic
    Do Until rsTBMERK.EOF
       txtCARI.AddItem rsTBMERK("kdbarang")
       rsTBMERK.MoveNext
    Loop
End Sub

Private Sub Form_Load()
txtCARI.Text = ""
End Sub


Public Function nvl(isi, kondisi)
    If IsNull(isi) = True Then
        nvl = kondisi
    Else
        nvl = isi
    End If
End Function



Private Sub Image1_Click()
    Query = "select tbstok.*,tbbarang.namabarang,tbbarang.stokakhir from tbstok,tbbarang where tbstok.kdbarang =tbbarang.kdbarang and tbbarang.kdbarang = '" & txtCARI.Text & "'"
     Set recordset = koneksi.Execute(Query, , adCmdText)
        If recordset.EOF Then
            MsgBox "TIDAK MENEMUKAN KODE BARANG! " _
            & " - " & txtCARI.Text & " - dalam tabel", _
            vbInformation, "Informasi"
            txtCARI.Text = ""
            txtCARI.SetFocus
        Else
          Call TampilGrid
          txtStok.Text = recordset.Fields("stokakhir")
          Text1.Text = recordset.Fields("namaBarang")
        End If
End Sub

Private Sub TbCetak_Click()
If txtCARI.Text = "" Then
Else
    Dim SQL1 As String
        SQL1 = ""
        SQL1 = "select tbstok.*,tbbarang.namabarang,tbbarang.stokakhir,tbmerk.nama as merk,tbjenis.nama as jenis from tbstok,tbbarang,tbmerk,tbjenis where tbstok.kdbarang =tbbarang.kdbarang and tbbarang.kdbarang = '" & txtCARI.Text & "' AND tbbarang.idmerk=tbmerk.idmerk AND tbbarang.idjenis=tbjenis.idjenis"
            Set rs_BARANG = koneksi.Execute(SQL1)
            If rs_BARANG.BOF Then
                MsgBox "DATA BARANG TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                "Informasi"
            Else
                 With Me.CRSTOK
                    .ReportFileName = App.Path & "\Report\tbstok.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    .SelectionFormula = "{tbbarang.kdbarang}='" & txtCARI.Text & "'"
                    .Action = 1
                End With
            End If
    End If
End Sub

Private Sub TbRefresh_Click()
    txtCARI.Text = ""
    txtStok.Text = ""
    Text1.Text = ""
    FGBARANG.Clear
    BARIS = 1
End Sub

Private Sub txtCARI_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Sub TampilGrid()
    Dim BARIS As Integer
    
    FGBARANG.Clear
    Call AktifGridBarang
     
         
    FGBARANG.Rows = 2
    BARIS = 0
     
     
   Set rs_BARANG = New ADODB.recordset
   SQL = "select tbstok.*,tbbarang.namabarang,tbbarang.stokakhir from tbstok,tbbarang where tbstok.kdbarang =tbbarang.kdbarang and tbbarang.kdbarang = '" & txtCARI.Text & "'"
    Set rs_BARANG = koneksi.Execute(SQL, , adCmdText)
   
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
            FGBARANG.TextMatrix(BARIS, 1) = ""
            FGBARANG.TextMatrix(BARIS, 2) = ""
            FGBARANG.TextMatrix(BARIS, 3) = "+" & .Fields("masuk")
            FGBARANG.TextMatrix(BARIS, 4) = "-" & .Fields("keluar")
            FGBARANG.TextMatrix(BARIS, 5) = .Fields("stok")
            FGBARANG.TextMatrix(BARIS, 6) = .Fields("no_bukti")
            FGBARANG.TextMatrix(BARIS, 7) = .Fields("keterangan")
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
        .ColWidth(1) = 0
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 2
        .Row = 0
        .Text = "NAMA BARANG"
        .CellFontBold = True
        .ColWidth(2) = 0
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 3
        .Row = 0
        .Text = "QTY IN"
        .CellFontBold = True
        .ColWidth(3) = 800
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 4
        .Row = 0
        .Text = "QTY OUT"
        .CellFontBold = True
        .ColWidth(4) = 800
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 5
        .Row = 0
        .Text = "STOK"
        .CellFontBold = True
        .ColWidth(5) = 700
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 6
        .Row = 0
        .Text = "NO BUKTI"
        .CellFontBold = True
        .ColWidth(6) = 3000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 7
        .Row = 0
        .Text = "KETERANGAN"
        .CellFontBold = True
        .ColWidth(7) = 3000
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
