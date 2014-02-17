VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form_LihatReturDis 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Data Barang"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   14865
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
   Icon            =   "Form_LihatReturDis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   14865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Retur Distributor"
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
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
         Left            =   8640
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
      Begin VB.Image TbRefresh 
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Left            =   2280
         Picture         =   "Form_LihatReturDis.frx":57E2
         Stretch         =   -1  'True
         ToolTipText     =   "Refresh"
         Top             =   360
         Width           =   510
      End
      Begin VB.Image TbHapus 
         BorderStyle     =   1  'Fixed Single
         Height          =   480
         Left            =   1560
         Picture         =   "Form_LihatReturDis.frx":6426
         Stretch         =   -1  'True
         ToolTipText     =   "Konfirm"
         Top             =   360
         Width           =   480
      End
      Begin VB.Image tbTambah 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   240
         Picture         =   "Form_LihatReturDis.frx":6D6A
         Stretch         =   -1  'True
         ToolTipText     =   "Tambah"
         Top             =   360
         Width           =   480
      End
      Begin VB.Image TbUbah 
         BorderStyle     =   1  'Fixed Single
         Height          =   480
         Left            =   840
         Picture         =   "Form_LihatReturDis.frx":72F2
         Stretch         =   -1  'True
         ToolTipText     =   "Ubah Data"
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   555
         Left            =   13920
         Picture         =   "Form_LihatReturDis.frx":7DB6
         ToolTipText     =   "Refresh"
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.TextBox txtkode 
      Enabled         =   0   'False
      Height          =   390
      Left            =   2160
      TabIndex        =   0
      Top             =   1440
      Width           =   2775
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGBARANG 
      Height          =   6135
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   10821
      _Version        =   393216
      Cols            =   7
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
      _Band(0).Cols   =   7
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "No Bukti Retur"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1500
   End
End
Attribute VB_Name = "Form_LihatReturDis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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

Private Sub Form_Load()
    If Form_utama.StatusBar1.Panels(2).Text = "ADMIN" Then
        TbHapus.Enabled = False
    Else
        TbHapus.Enabled = True
    End If
    cmbcari.AddItem "Cari Semua"
    cmbcari.AddItem "No Bukti Retur"
    cmbcari.AddItem "Kode Distributor"
    Call tampilgrid
End Sub
Sub tampilgrid()
    Dim BARIS As Integer
    
    FGBARANG.Clear
    Call AktifGridBarang
     
         
    FGBARANG.Rows = 2
    BARIS = 0
     
     
   Set rs_retur = New ADODB.recordset
   If cmbcari.Text = "No Bukti Retur" Then
        Query = "call ReturKode('%" & txtcari.Text & "%')"
   ElseIf cmbcari.Text = "Kode Distributor" Then
        Query = "call ReturDis('%" & txtcari.Text & "%')"
   Else
        Query = "call TampilRetur()"
   End If
   Set rs_retur = koneksi.Execute(Query, , adCmdText)
   
     If rs_retur.EOF Then
         MsgBox "DATA KOSONG!", _
         vbInformation + vbOKOnly, "Informasi"
         Exit Sub
     Else
         With rs_retur
            .MoveFirst
         Do While Not .EOF
            
            BARIS = BARIS + 1
            FGBARANG.Rows = BARIS + 1
            FGBARANG.TextMatrix(BARIS, 0) = BARIS
            FGBARANG.TextMatrix(BARIS, 1) = nvl(.Fields("kdReturSup"), "0")
            FGBARANG.TextMatrix(BARIS, 2) = nvl(Format(.Fields("tglretur"), "yyyy-dd-mm"), "0")
            FGBARANG.TextMatrix(BARIS, 3) = nvl(.Fields("namaDistributor"), "0")
            FGBARANG.TextMatrix(BARIS, 5) = nvl(.Fields("user_ubah"), "0")
            FGBARANG.TextMatrix(BARIS, 6) = nvl(.Fields("tgl_ubah"), "0")
            FGBARANG.TextMatrix(BARIS, 4) = nvl(.Fields("kddistributor"), "0")
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
        .Text = "KODE DISTRIBUTOR"
        .CellFontBold = True
        .ColWidth(4) = 2000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 5
        .Row = 0
        .Text = "USER UBAH"
        .CellFontBold = True
        .ColWidth(5) = 2000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 6
        .Row = 0
        .Text = "TGL UBAH"
        .CellFontBold = True
        .ColWidth(6) = 2000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
    End With
End Sub

Private Sub TbHapus_Click()
        If txtkode.Text = "" Then
        MsgBox "PILIH NO BUKTI TERIMA!", _
         vbInformation + vbOKOnly, "Informasi"
         Exit Sub
    Else
        Pesan = MsgBox("Yakin Akan Dikonfirm???" _
            , vbQuestion + vbYesNo, "Konfirmasi")
        If Pesan = vbYes Then
                SQL1 = "select * FROM tbtmpretursup1 where kdretursup='" & txtkode & "'"
                Set rs_retur = koneksi.Execute(SQL1, , adCmdText)
                
                
                Query = "CALL TambahRetsup('" & txtkode & "','" & Format(rs_retur.Fields(1), "yyyy-mm-dd") & "','" & rs_retur.Fields(2) & "','" & Form_utama.StatusBar1.Panels(1).Text & "',now(),'Y')"
                    koneksi.Execute Query, , adCmdText
                    
                    SQL = "Select *From tbtmpretursup where kdRetursup='" & txtkode & "'"
                    Set recordsett = koneksi.Execute(SQL, , adCmdText)
                    recordsett.MoveFirst
            
            
                While Not recordsett.EOF
                    a = recordsett.Fields(1)
                
                    SQL2 = "Select * From tbBARANG where KDbarang ='" & a & "'"
                    Set rs_BARANG = koneksi.Execute(SQL2, , adCmdText)
               
                    kurang = rs_BARANG.Fields(10) - recordsett.Fields(3)
                    SQL3 = "update tbBARANG set StokAkhir='" & kurang & "' Where KDbarang ='" & a & "'"
                    Set recordset = koneksi.Execute(SQL3, , adCmdText)
                    
                    sql4 = "insert into tbstok(kdBarang,no_bukti,masuk,keluar,stok,user_ubah,tgl_ubah,keterangan) values('" & a & "','" & txtkode & "','0','" & recordsett.Fields(4) & "','" & kurang & "','" & Form_utama.StatusBar1.Panels(1).Text & "',now(),'RETUR DISTRIBUTOR')"
                    Set rs_STOK = koneksi.Execute(sql4, , adCmdText)
                    
                    recordsett.MoveNext
                Wend
                sql5 = "update tbtmpretursup1 set flag='Y' where kdretursup='" & txtkode & "'"
                Set rs_TERIMA = koneksi.Execute(sql5, , adCmdText)
                
                MsgBox "DATA BERHASIL DISIMPAN" + Chr(13) + "NOTE:", 64, "Konfirmasi"
                'Call Form_Activate
                Call Form_Load
                txtkode.Text = ""
        End If
    End If
End Sub

Private Sub TbRefresh_Click()
    Call tampilgrid
    Me.FGBARANG.Refresh
    txtcari.Text = ""
    txtkode.Text = ""
End Sub

Private Sub tbTambah_Click()
    txtkode.Text = ""
    Form_returDis.Show 1

End Sub

Private Sub TbUbah_Click()
    If txtkode.Text = "" Then
        MsgBox "PILIH NO BUKTI TERIMA!", _
         vbInformation + vbOKOnly, "Informasi"
         Exit Sub
    Else
        Form_ubahReturDis.Show 1
    End If
End Sub

Private Sub txtcari_Change()
  If cmbcari.Text = "No Bukti Retur" Then
        Query = "call ReturKode('%" & txtcari.Text & "%')"
   ElseIf cmbcari.Text = "Kode Distributor" Then
        Query = "call ReturDis('%" & txtcari.Text & "%')"
   Else
        Query = "call TampilRetur()"
   End If
        Set recordset = koneksi.Execute(Query, , adCmdText)
        If recordset.EOF Then
            MsgBox "TIDAK MENEMUKAN DATA! " _
            & " - " & txtcari.Text & " - dalam tabel", _
            vbInformation, "Informasi"
            
            txtcari.Text = ""
            txtcari.SetFocus
        Else
          Call tampilgrid
        End If
End Sub
Public Function nvl(isi, kondisi)
    If IsNull(isi) = True Then
        nvl = kondisi
    Else
        nvl = isi
    End If
End Function
