VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form_UbahReturUn 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Data Barang"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   13575
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
   Icon            =   "Form_UbahReturUn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   13575
   StartUpPosition =   2  'CenterScreen
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
      Left            =   12240
      Picture         =   "Form_UbahReturUn.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Frame Frame_iden 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Identifikasi"
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   13215
      Begin VB.TextBox txtkodeBukti 
         Height          =   390
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txttglbukti 
         Height          =   390
         Left            =   2640
         TabIndex        =   17
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtDist 
         Height          =   390
         Left            =   2640
         TabIndex        =   16
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtkodeUnit 
         Height          =   390
         Left            =   9360
         TabIndex        =   15
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Unit"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   25
         Top             =   840
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
         TabIndex        =   24
         Top             =   360
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
         TabIndex        =   23
         Top             =   360
         Width           =   60
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti Retur"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   6960
         TabIndex        =   22
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   8880
         TabIndex        =   21
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
         Left            =   8880
         TabIndex        =   20
         Top             =   840
         Width           =   60
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Unit"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   6960
         TabIndex        =   19
         Top             =   840
         Width           =   1035
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2280
         Picture         =   "Form_UbahReturUn.frx":5E77
         Stretch         =   -1  'True
         ToolTipText     =   "Find"
         Top             =   840
         Width           =   480
      End
   End
   Begin VB.Frame Frame_isi 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Data Barang Retur"
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   13215
      Begin VB.TextBox txtmerk 
         Height          =   390
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtalasan 
         Height          =   390
         Left            =   8760
         TabIndex        =   7
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox txtJml 
         Height          =   390
         Left            =   7680
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtNamaBa 
         Height          =   390
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtKodeBa 
         Height          =   390
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Merk"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   5760
         TabIndex        =   13
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Alasan Retur"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   10200
         TabIndex        =   12
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Retur"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   7560
         TabIndex        =   11
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   720
         TabIndex        =   10
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3240
         TabIndex        =   9
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.Frame pop1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   360
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   2655
      Begin MSComctlLib.ListView lvnm1 
         Height          =   3735
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   2535
         _ExtentX        =   4471
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
         TabIndex        =   2
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
   Begin VB.Frame Frame_ret 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Barang Yang DiRetur"
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   240
      TabIndex        =   27
      Top             =   4080
      Width           =   13215
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGRETUR 
         Height          =   1575
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   2778
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
   End
   Begin VB.TextBox txtbukti 
      Height          =   630
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image TbUbah 
      Height          =   480
      Left            =   12240
      Picture         =   "Form_UbahReturUn.frx":6181
      Stretch         =   -1  'True
      ToolTipText     =   "Ubah"
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image TbHapus 
      Height          =   480
      Left            =   12840
      Picture         =   "Form_UbahReturUn.frx":6C45
      ToolTipText     =   "Hapus"
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image tbTambah 
      Height          =   495
      Left            =   11640
      Picture         =   "Form_UbahReturUn.frx":7889
      Stretch         =   -1  'True
      ToolTipText     =   "Tambah"
      Top             =   3360
      Width           =   480
   End
End
Attribute VB_Name = "Form_UbahReturUn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim jml As Integer
    Dim hrg As Double
    Dim BARIS As Integer
    Dim stok As Integer
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Const LVM_FIRST = &H1000
Const LVM_SCROLL = (LVM_FIRST + 20)
Public Function nvl(isi, kondisi)
    If IsNull(isi) = True Then
        nvl = kondisi
    Else
        nvl = isi
    End If
End Function



Private Sub cmdKELUAR_Click()
    If BARIS <= 0 Then
MsgBox "BARANG MASIH KOSONG ", _
        vbOKOnly + vbCritical, "Konfirmasi"
Else
    Unload Me
End If
End Sub




Private Sub tdkAktif()
    
    frame_iden.Enabled = False
    
    cmdKELUAR.Enabled = True
End Sub

Private Sub bersih()
    txtDist.Text = ""
    txtkodeBukti.Text = ""
    txtkodeUnit.Text = ""
End Sub
Public Sub RecTerakhir()
Dim Query As String
On Error Resume Next
    Query = "select max(kdretursup) from tbTMPretursup1"
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
        Select Case NOBM
            Case 0 To 9
               Me.txtkodeBukti.Text = "000" & Trim(Str(NOBM)) + "/" + "RS" + "/" + Mid(txttglbukti.Text, 6, 2) + "/" + Right(Date, 2)
            Case 10 To 99
               Me.txtkodeBukti.Text = "00" & Trim(Str(NOBM)) + "/" + "RS" + "/" + Mid(txttglbukti.Text, 6, 2) + "/" + Right(Date, 2)
            Case 100 To 999
               Me.txtkodeBukti.Text = "0" & Trim(Str(NOBM)) + "/" + "RS" + "/" + Mid(txttglbukti.Text, 6, 2) + "/" + Right(Date, 2)
            Case 1000 To 9999
               Me.txtkodeBukti.Text = Trim(Str(NOBM)) + "/" + "RS" + "/" + Mid(txttglbukti.Text, 6, 2) + "/" + Right(Date, 2)
         End Select
    Else
       Me.txtkodeBukti.Text = "0001" + "/" + "RS" + "/" + Mid(txttglbukti.Text, 6, 2) + "/" + Right(Date, 2)
    End If
End Sub

Private Sub Form_Click()
    pop1.Visible = False
End Sub

Private Sub Form_Load()
    txttglbukti.Text = Format(Date, "yyyy-mm-dd")
    tdkAktif
    bersih
    kosong
    pop1.Visible = False
    ambil_nilai
End Sub

Private Sub frame_iden_Click()
    pop1.Visible = False
End Sub


Private Sub Aktif()
    'Frame_det.Enabled = True
    frame_iden.Enabled = True
    Frame_isi.Enabled = True
    Frame_ret.Enabled = True
    'Frame_terima.Enabled = True
    
    'cmdSIMPAN.Enabled = True
    'cmdTAMBAH.Enabled = False
    'cmdBATAL.Enabled = True
    cmdKELUAR.Enabled = False
    FGRETUR.Enabled = True
End Sub

Private Sub Frame_isi_Click()
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



Private Sub lvnm1_Click()

            If lvnm1.SelectedItem <> "" Then
                txtNamaBa.Text = lvnm1.SelectedItem
                Query = "call BarangNama('%" & txtNamaBa.Text & "%')"
                Set rs_BARANG = koneksi.Execute(Query, , adCmdText)
                If rs_BARANG.EOF Then
                    MsgBox "DATA TIDAK ADA" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
                Else
                    txtNamaBa.Text = nvl(rs_BARANG.Fields("namaBarang"), "0")
                    txtKodeBa.Text = nvl(rs_BARANG.Fields("kdBarang"), "0")
                    hrg = nvl(rs_BARANG.Fields("hargadasar"), "0")
                    stok = nvl(rs_BARANG.Fields("stokAkhir"), "0")
                    txtmerk.Text = nvl(rs_BARANG.Fields("merk"), "0")
                    pop1.Visible = False
                    'txtQty.SetFocus
                End If
                pop1.Visible = False
            End If

End Sub


Private Sub kosong()
    txtKodeBa.Text = ""
    txtNamaBa.Text = ""
    txtJml.Text = "0"
    txtalasan.Text = ""
    txtmerk.Text = ""
    jml = 0
    hrg = 0
End Sub

Private Sub LvNm_DblClick()
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


Private Sub TbHapus_Click()
 Query = "CALL HapusTmpRetUn('" & txtkodeBukti.Text & "','" & txtKodeBa.Text & "')"
        Pesan = MsgBox("Bener Mau Dihapus !" _
            , vbQuestion + vbYesNo, "Konfirmasi")
        If Pesan = vbYes Then
            Set recordset = koneksi.Execute(Query, , adCmdText)
            Call kosong
            Me.FGRETUR.Refresh
            Call TampilGridRetur
                TbUbah.Enabled = False
                TbHapus.Enabled = False
                tbTambah.Enabled = True
                BARIS = BARIS - 1
        End If
End Sub

Private Sub tbTambah_Click()
   If txtKodeBa.Text = "" Then
        MsgBox "KODE BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        txtKodeBa.SetFocus
    ElseIf txtJml.Text = "" Then
        MsgBox "JUMLAH BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        txtJml.SetFocus
    ElseIf txtalasan.Text = "" Then
        MsgBox "ALASAN TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        txtalasan.SetFocus
    ElseIf txtJml.Text > stok Then
        MsgBox "JUMLAH RETUR TIDAK BOLEH LEBIH BESAR DARI TERIMA BARANG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    Else
        Pesan = MsgBox("YAKIN MAU DIINPUT ??" _
            , vbQuestion + vbYesNo, "Konfirmasi")
        If Pesan = vbYes Then
            SQL = "select * from tbtmpreturUN where kdReturUN='" & txtkodeBukti & "' and kdBarang='" & txtKodeBa & "'"
            Set rs_BARANG = koneksi.Execute(SQL, , adCmdText)
            If Not rs_BARANG.EOF Then
                 MsgBox "DATA BARANG SUDAH ADA, TIDAK BISA DIINPUT LAGI!", _
                 vbInformation + vbOKOnly, "Informasi"
                 Exit Sub
            Else
                Dim TOTAL As Double
                TOTAL = Val(Int(txtJml.Text)) * Val(Int(hrg))
                Query = "call TambahTmpRetUN('" & txtkodeBukti & "','" & txtKodeBa & "','" & Int(hrg) & "','" & txtJml & "','" & Int(TOTAL) & "','" & txtalasan & "')"
                koneksi.Execute Query, , adCmdText
                Call kosong
                Me.FGRETUR.Refresh
                Call TampilGridRetur
                BARIS = BARIS + 1
            End If
        End If
    End If
End Sub
Sub TampilGridRetur()
    
    
    FGRETUR.Clear
    Call AktifGridRetur
     
         
    FGRETUR.Rows = 2
    BARIS = 0
     
     
   Set rs_retur = New ADODB.recordset
   Query = "Select tbtmpreturun.*,tbbarang.namaBarang from tbtmpreturun,tbbarang where tbbarang.KdBarang=tbtmpreturun.kdBarang and tbtmpreturun.kdReturun='" & txtkodeBukti & "'"
   Set rs_retur = koneksi.Execute(Query, , adCmdText)
   
     If rs_retur.EOF Then
     Else
         With rs_retur
            .MoveFirst
         Do While Not .EOF
            BARIS = BARIS + 1
            FGRETUR.Rows = BARIS + 1
            FGRETUR.TextMatrix(BARIS, 0) = BARIS
            FGRETUR.TextMatrix(BARIS, 1) = .Fields("kdBarang")
            FGRETUR.TextMatrix(BARIS, 2) = .Fields("namaBarang")
            FGRETUR.TextMatrix(BARIS, 3) = .Fields("jml")
            FGRETUR.TextMatrix(BARIS, 4) = Format(.Fields("harga"), "###,###,##0")
            FGRETUR.TextMatrix(BARIS, 5) = Format(.Fields("total"), "###,###,##0")
            FGRETUR.TextMatrix(BARIS, 6) = .Fields("alasan")
         .MoveNext
         Loop
         End With
     End If
End Sub
Private Sub AktifGridRetur()
    With FGRETUR
        
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
        .Text = "HARGA"
        .CellFontBold = True
        .ColWidth(4) = 2500
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 5
        .Row = 0
        .Text = "TOTAL"
        .CellFontBold = True
        .ColWidth(5) = 3000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 6
        .Row = 0
        .Text = "ALASAN"
        .CellFontBold = True
        .ColWidth(6) = 3000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
    End With
End Sub


Private Sub TbUbah_Click()
   If txtKodeBa.Text = "" Then
        MsgBox "KODE BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        txtKodeBa.SetFocus
    ElseIf txtJml.Text = "" Then
        MsgBox "JUMLAH BARANG TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        txtJml.SetFocus
    ElseIf txtalasan.Text = "" Then
        MsgBox "ALASAN TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
        txtalasan.SetFocus
    
    Else
        Pesan = MsgBox("YAKIN MAU DIUBAH ??" _
            , vbQuestion + vbYesNo, "Konfirmasi")
        If Pesan = vbYes Then
        
                Dim TOTAL As Double
                TOTAL = Val(Int(txtJml.Text)) * Val(Int(hrg))
                Query = "call EditTmpRetun('" & txtkodeBukti & "','" & txtKodeBa & "','" & txtJml & "','" & Int(hrg) & "','" & Int(TOTAL) & "','" & txtalasan & "')"
                koneksi.Execute Query, , adCmdText
                Call kosong
                Me.FGRETUR.Refresh
                Call TampilGridRetur
        End If
    End If
End Sub

Private Sub txtalasan_KeyPress(KeyAscii As Integer)
Call BlokKarakter(KeyAscii)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub FGRETUR_DblClick()
If BARIS <= 0 Then
Else
    Dim barisGrid As String
    barisGrid = FGRETUR.Row
    
    If FGRETUR.Rows <> 1 Then
        txtKodeBa.Text = _
            FGRETUR.TextMatrix(barisGrid, 1)
        txtNamaBa.Text = _
            FGRETUR.TextMatrix(barisGrid, 2)
        txtJml.Text = _
            FGRETUR.TextMatrix(barisGrid, 3)
        hrg = _
            FGRETUR.TextMatrix(barisGrid, 4)
        txtalasan.Text = _
            FGRETUR.TextMatrix(barisGrid, 6)
        jml = _
            FGRETUR.TextMatrix(barisGrid, 3)
    Else
        Exit Sub
    End If
        
    TbUbah.Enabled = True
    TbHapus.Enabled = True
    tbTambah.Enabled = False
End If

End Sub
Function ListViewScroll(lvnm1 As ListView, ByVal dx As Long, ByVal dy As Long)
    SendMessage lvnm1.hwnd, LVM_SCROLL, dx, ByVal dy
    SendMessage LvNm.hwnd, LVM_SCROLL, dx, ByVal dy
End Function

Private Sub txtJml_KeyPress(KeyAscii As Integer)
    Call HanyaNomor(KeyAscii)
    If KeyAscii = 13 Then
        txtalasan.SetFocus
    End If
End Sub

Private Sub txtNamaBa_Change()
    pop1.Visible = True
    txtNamaBa.SelStart = Len(txtNamaBa.Text)
    namaBarang
End Sub
Private Sub namaBarang()
   Query = "select tbbarang.*,tbmerk.nama" _
            & " From tbBarang, tbmerk, tbkirimUn, tbdetKIRIMUn" _
            & " Where tbBarang.idmerk = tbmerk.idmerk" _
            & " and tbkirimUn.kdkirimUn=tbdetkirimun.kdkirimun" _
            & " and tbdetkirimUn.kdbarang=tbbarang.kdbarang" _
            & " and tbkirimun.kdunit='" & txtkodeUnit & "'" _
            & " group by tbbarang.kdbarang"
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

Private Sub ambil_nilai()
    Dim kode As String
    kode = Form_LapReturUnit.txtkode.Text
    
    Query = "Select tbtmpreturun1.*,tbunit.namaunit " _
            & " from tbtmpreturun1,tbunit " _
            & " where tbtmpreturun1.kdUnit=tbunit.kdunit " _
            & "  and tbtmpreturun1.kdreturun='" & kode & "'"
            Set rs_TERIMA = koneksi.Execute(Query)
      If Not rs_TERIMA.EOF Then
         txttglbukti.Text = nvl(Format(rs_TERIMA("tglreturUn"), "yyyy-mm-dd"), "")
         txtkodeUnit.Text = nvl(rs_TERIMA("kdUnit"), "")
         txtkodeBukti.Text = nvl(rs_TERIMA("kdreturUn"), "")
         txtDist.Text = nvl(rs_TERIMA("namaUnit"), "")
    End If
    frame_iden.Enabled = False
    Call TampilGridRetur
End Sub


