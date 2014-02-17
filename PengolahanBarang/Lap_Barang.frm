VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Lap_Barang 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Stok Barang"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4995
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
   Icon            =   "Lap_Barang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4680
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Laporan Data Barang"
      ForeColor       =   &H00000000&
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4455
      Begin VB.TextBox txtcari 
         Height          =   390
         Left            =   840
         TabIndex        =   3
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   390
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   3855
      End
      Begin VB.CommandButton cmdlihat 
         Caption         =   "&CETAK"
         Height          =   975
         Left            =   1560
         Picture         =   "Lap_Barang.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Lap_Barang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdlihat_Click()
        If Combo1.Text = "--PILIH BERDASARKAN--" Then
        MsgBox "PILIH BERDASARKAN APA??! ", _
            vbInformation, "Informasi"
            Combo1.SetFocus
    ElseIf Combo1.Text = "SELURUH DATA BARANG" Then
        Dim SQL As String
        SQL = ""
        SQL = "Call TampilBarang()"
            Set rs_BARANG = koneksi.Execute(SQL)
            If rs_BARANG.BOF Then
                MsgBox "DATA BARANG TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                "Informasi"
                txtcari.SetFocus
            Else
                With Me.CrystalReport1
                    .ReportFileName = App.Path & "\Report\tbBarang.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    .Action = 1
                End With
            End If
    ElseIf Combo1.Text = "JENIS BARANG" Then
        Dim SQL1 As String
        SQL1 = ""
        SQL1 = "SELECT tbjenis.nama as jenis,tbmerk.nama as merk,tbbarang.* From tbBarang, tbjenis, tbmerk Where tbBarang.idjenis = tbjenis.idjenis AND tbbarang.idmerk=tbmerk.idmerk AND tbjenis.nama ='" & txtcari & "'"
            Set rs_BARANG = koneksi.Execute(SQL1)
            If rs_BARANG.BOF Then
                MsgBox "DATA BARANG TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                "Informasi"
                txtcari.SetFocus
            Else
                With Me.CrystalReport1
                    .ReportFileName = App.Path & "\Report\tbBarangJenis.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    .SelectionFormula = "{tbjenis.nama}='" & txtcari.Text & "'"
                    .Action = 1
                End With
            End If
        ElseIf Combo1.Text = "MERK BARANG" Then
        Dim SQL3 As String
        SQL3 = ""
        SQL3 = "SELECT tbjenis.nama as jenis,tbmerk.nama as merk,tbbarang.* From tbBarang, tbjenis, tbmerk Where tbBarang.idjenis = tbjenis.idjenis AND tbbarang.idmerk=tbmerk.idmerk AND tbmerk.nama ='" & txtcari & "'"
            Set rs_BARANG = koneksi.Execute(SQL3)
            If rs_BARANG.BOF Then
                MsgBox "DATA BARANG TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                "Informasi"
                txtcari.SetFocus
            Else
                DataEnvironment1.Commands(3).CommandText = SQL3
        
            With Me.CrystalReport1
                    .ReportFileName = App.Path & "\Report\tbBarangMerk.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    .SelectionFormula = "{tbmerk.nama}='" & txtcari.Text & "'"
                    .Action = 1
                End With
            End If
    End If
    
End Sub

Private Sub Form_Load()
Combo1.Text = "--PILIH BERDASARKAN--"
    Combo1.AddItem "SELURUH DATA BARANG"
    Combo1.AddItem "JENIS BARANG"
    Combo1.AddItem "MERK BARANG"
End Sub
