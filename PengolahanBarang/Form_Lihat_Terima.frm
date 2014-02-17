VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form_Lihat_Terima 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Data Barang"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   14850
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
   Icon            =   "Form_Lihat_Terima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   14850
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtkode 
      Enabled         =   0   'False
      Height          =   390
      Left            =   2160
      TabIndex        =   5
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Terima Barang"
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   14655
      Begin VB.ComboBox cmbcari 
         Height          =   390
         Left            =   8640
         TabIndex        =   3
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtcari 
         Height          =   390
         Left            =   11160
         TabIndex        =   2
         Top             =   480
         Width           =   2775
      End
      Begin VB.Image TbUbah 
         BorderStyle     =   1  'Fixed Single
         Height          =   480
         Left            =   840
         Picture         =   "Form_Lihat_Terima.frx":57E2
         Stretch         =   -1  'True
         ToolTipText     =   "Ubah Data"
         Top             =   480
         Width           =   480
      End
      Begin VB.Image tbTambah 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   240
         Picture         =   "Form_Lihat_Terima.frx":62A6
         Stretch         =   -1  'True
         ToolTipText     =   "Tambah"
         Top             =   480
         Width           =   480
      End
      Begin VB.Image TbHapus 
         BorderStyle     =   1  'Fixed Single
         Height          =   480
         Left            =   1560
         Picture         =   "Form_Lihat_Terima.frx":682E
         Stretch         =   -1  'True
         ToolTipText     =   "Konfirm"
         Top             =   480
         Width           =   480
      End
      Begin VB.Image TbRefresh 
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Left            =   2280
         Picture         =   "Form_Lihat_Terima.frx":7172
         Stretch         =   -1  'True
         ToolTipText     =   "Refresh"
         Top             =   480
         Width           =   510
      End
      Begin VB.Image Image1 
         Height          =   555
         Left            =   13920
         Picture         =   "Form_Lihat_Terima.frx":7DB6
         ToolTipText     =   "Refresh"
         Top             =   360
         Width           =   555
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGBARANG 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   10821
      _Version        =   393216
      Cols            =   6
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
      _Band(0).Cols   =   6
   End
   Begin Crystal.CrystalReport crTerimaKode 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "No Bukti Terima"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1680
   End
End
Attribute VB_Name = "Form_Lihat_Terima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbcari_Click()
    txtcari.SetFocus
End Sub



Private Sub FGBARANG_DblClick()
    Dim barisGrid As String
    barisGrid = FGBARANG.Row
    
    If FGBARANG.Rows <> 1 Then
        txtkode.Text = _
            FGBARANG.TextMatrix(barisGrid, 1)
    Else
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
    If Form_utama.StatusBar1.Panels(2).Text = "ADMIN" Then
        TbHapus.Enabled = False
    Else
        TbHapus.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    cmbcari.AddItem "Kode Bukti Terima"
    cmbcari.AddItem "Kode Distributor"
    Call TampilGrid
    txtkode.Text = ""
    
    If Form_utama.StatusBar1.Panels(2).Text = "ADMIN" Then
        TbHapus.Enabled = False
    Else
        TbHapus.Enabled = True
    End If
End Sub



Private Sub TbHapus_Click()
    If txtkode.Text = "" Then
        MsgBox "PILIH NO BUKTI TERIMA!", _
         vbInformation + vbOKOnly, "Informasi"
         Exit Sub
    Else
    Pesan = MsgBox("Yakin Form Penerimaan Barang Akan DiKonfirm?", vbYesNo + vbQuestion, "Konfirmasi Konfirm")
        If Pesan = 6 Then
        SQL1 = "select * FROM tbtmpterima1 where kdkirim='" & txtkode & "'"
            Set rs_TERIMA1 = koneksi.Execute(SQL1, , adCmdText)
        
        Query = "call TambahTerima('" & rs_TERIMA1.Fields(0) & "','" & Format(rs_TERIMA1.Fields(1), "yyyy-mm-dd") & "','" & rs_TERIMA1.Fields(2) & "','" & Format(rs_TERIMA1.Fields(3), "yyyy-mm-dd") & "','" & rs_TERIMA1.Fields(4) & "','" & Form_utama.StatusBar1.Panels(1).Text & "',now(),'Y')"
        koneksi.Execute Query, , adCmdText
        
        SQL = "Select * From tbtmpterima where kdKirim='" & txtkode & "'"
        Set recordsett = koneksi.Execute(SQL, , adCmdText)
        recordsett.MoveFirst

    
    While Not recordsett.EOF
        a = recordsett.Fields(1)
    
        SQL2 = "Select *From tbBARANG where KDbarang ='" & a & "'"
        Set rs_BARANG = koneksi.Execute(SQL2, , adCmdText)
   
        kurang = rs_BARANG.Fields(10) + recordsett.Fields(2)
        SQL3 = "update tbBARANG set StokAkhir='" & kurang & "',HargaDasar='" & recordsett.Fields(3) & "',hargafixed='" & recordsett.Fields(5) & "' Where KDbarang ='" & a & "'"
        Set recordset = koneksi.Execute(SQL3, , adCmdText)
        
        sql4 = "insert into tbstok(kdBarang,no_bukti,masuk,keluar,stok,user_ubah,tgl_ubah,keterangan)" _
        & " values('" & a & "','" & txtkode & "','" & recordsett.Fields(2) & "','0','" & kurang & "','" & Form_utama.StatusBar1.Panels(1).Text & "',now(),'TERIMA DISTRIBUTOR')"
        Set rs_STOK = koneksi.Execute(sql4, , adCmdText)
        recordsett.MoveNext
        
    Wend
    sql5 = "update tbtmpterima1 set flag='Y' where kdkirim='" & txtkode & "'"
    Set rs_TERIMA = koneksi.Execute(sql5, , adCmdText)
    
    MsgBox "DATA BERHASIL DIKONFIRM, SILAHKAN CETAK NOTA" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        Pesan = MsgBox("Cetak Form Penerimaan Barang?", vbYesNo + vbQuestion, "Konfirmasi Cetak")
            If Pesan = 6 Then
                
                SQL1 = ""
                SQL1 = "select kdkirim from tbterima where konfirm='Y' and kdkirim='" & txtkode & "' order by kdkirim"
                    Set rs_BARANG = koneksi.Execute(SQL1)
                    If rs_BARANG.BOF Then
                        MsgBox "DATA TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                        "Informasi"
                        
                    Else
                        With Me.crTerimaKode
                            .ReportFileName = App.Path & "\Report\tbTerima.rpt"
                            .WindowState = crptMaximized
                            .RetrieveDataFiles
                            .SelectionFormula = "{tbterima.kdkirim}='" & txtkode.Text & "'"
                            .Action = 1
                        End With
                    End If
            End If
    Call Form_Load
    End If
    End If
End Sub

Private Sub TbRefresh_Click()
    Call TampilGrid
    Me.FGBARANG.Refresh
    txtcari.Text = ""
    txtkode.Text = ""
End Sub

Private Sub tbTambah_Click()
    txtkode.Text = ""
    Form_TerimaBarang.Show 1
End Sub
Sub TampilGrid()
    Dim BARIS As Integer
    
    FGBARANG.Clear
    Call AktifGridBarang
     
         
    FGBARANG.Rows = 2
    BARIS = 0
     
     
   Set rs_BARANG = New ADODB.recordset
   If cmbcari.Text = "Kode Bukti Terima" Then
        Query = "call TerimaKode('%" & txtcari.Text & "%')"
   ElseIf cmbcari.Text = "Kode Distributor" Then
        Query = "call TerimaDis('%" & txtcari.Text & "%')"
   Else
        Query = "call TerimaTerima()"
   End If
   Set rs_TERIMA = koneksi.Execute(Query, , adCmdText)
   
     If rs_TERIMA.EOF Then
         MsgBox "DATA KOSONG!", _
         vbInformation + vbOKOnly, "Informasi"
         Exit Sub
     Else
         With rs_TERIMA
            .MoveFirst
         Do While Not .EOF
            
            BARIS = BARIS + 1
            FGBARANG.Rows = BARIS + 1
            FGBARANG.TextMatrix(BARIS, 0) = BARIS
            FGBARANG.TextMatrix(BARIS, 1) = nvl(.Fields("kdkirim"), "0")
            FGBARANG.TextMatrix(BARIS, 2) = nvl(.Fields("tgl_terima"), "0")
            FGBARANG.TextMatrix(BARIS, 3) = nvl(.Fields("namaDistributor"), "0")
            FGBARANG.TextMatrix(BARIS, 4) = nvl(.Fields("user_ubah"), "0")
            FGBARANG.TextMatrix(BARIS, 5) = nvl(.Fields("tgl_ubah"), "0")
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
        .Text = "NOMOR BUKTI"
        .CellFontBold = True
        .ColWidth(1) = 1500
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 2
        .Row = 0
        .Text = "TANGGAL"
        .CellFontBold = True
        .ColWidth(2) = 2000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 3
        .Row = 0
        .Text = "NAMA DISTRIBUTOR"
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
        
    End With
End Sub

Private Sub TbUbah_Click()
    If txtkode.Text = "" Then
        MsgBox "PILIH NO BUKTI TERIMA!", _
         vbInformation + vbOKOnly, "Informasi"
         Exit Sub
    Else
        Form_ubahTerima.Show 1
    End If
End Sub

Private Sub txtcari_Change()
       If cmbcari.Text = "Kode Bukti Terima" Then
        Query = "call TerimaKode('%" & txtcari.Text & "%')"
   ElseIf cmbcari.Text = "Kode Distributor" Then
        Query = "call TerimaDis('%" & txtcari.Text & "%')"
   Else
        Query = "call TerimaTerima()"
   End If
        Set recordset = koneksi.Execute(Query, , adCmdText)
        If recordset.EOF Then
            MsgBox "TIDAK MENEMUKAN DATA! " _
            & " - " & txtcari.Text & " - dalam tabel", _
            vbInformation, "Informasi"
            
            txtcari.Text = ""
            txtcari.SetFocus
        Else
          Call TampilGrid
        End If
End Sub
Public Function nvl(isi, kondisi)
    If IsNull(isi) = True Then
        nvl = kondisi
    Else
        nvl = isi
    End If
End Function
