VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form_MasterJenis 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Data Barang"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_MasterJenis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame pop1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   2760
      TabIndex        =   18
      Top             =   1560
      Visible         =   0   'False
      Width           =   3855
      Begin MSComctlLib.ListView lvnm1 
         Height          =   3855
         Left            =   0
         TabIndex        =   19
         Top             =   120
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   6800
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
            Text            =   "Nama Jenis"
            Object.Width           =   7056
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Halaman Input Jenis Barang"
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   7335
      Begin VB.TextBox txtkode 
         Height          =   390
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtnama 
         Height          =   390
         Left            =   2640
         TabIndex        =   11
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kode Jenis"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   360
         TabIndex        =   16
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   15
         Top             =   600
         Width           =   60
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nama Jenis"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   13
         Top             =   1080
         Width           =   60
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   7335
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
         Left            =   3720
         Picture         =   "Form_MasterJenis.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   120
         Picture         =   "Form_MasterJenis.frx":5E49
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   1320
         Picture         =   "Form_MasterJenis.frx":646B
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   2520
         Picture         =   "Form_MasterJenis.frx":6A8D
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   4920
         Picture         =   "Form_MasterJenis.frx":70FF
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   6120
         Picture         =   "Form_MasterJenis.frx":7734
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   6600
      Width           =   7335
      Begin VB.TextBox txtCARI 
         Height          =   390
         Left            =   1920
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Masukkan Nama"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1755
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGmerk 
      Height          =   2655
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   3
      GridColorFixed  =   255
      Appearance      =   0
      BandDisplay     =   1
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
      _Band(0).Cols   =   3
   End
End
Attribute VB_Name = "Form_MasterJenis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBATAL_Click()
    tdkAktifkan
    BERSIHKAN
End Sub

Private Sub cmdHAPUS_Click()
SQL = "CALL CekJenis('" & txtkode.Text & "')"
Set rsTBJENIS = koneksi.Execute(SQL, , adCmdText)
If Not rsTBJENIS.EOF Then
 MsgBox "DATA TIDAK DAPAT DIHAPUS, SEDANG DIGUNAKAN DI TABEL LAIN " + Chr(13) + "NOTE:", 64, "Konfirmasi"
 Call Form_Activate
Else
Query = "CALL HapusJenis('" & txtkode.Text & "')"
    Pesan = MsgBox("Bener Neeh Mau Dihapus !" _
            , vbQuestion + vbYesNo, "Konfirmasi")
    If Pesan = vbYes Then
       Set recordset = koneksi.Execute(Query, , adCmdText)
       Call Form_Activate
       Me.FGmerk.Refresh
    End If
End If
End Sub

Private Sub cmdKELUAR_Click()
Unload Me
End Sub

Private Sub cmdSIMPAN_Click()
If txtkode.Text = "" Then
    MsgBox "KODE JENIS TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    txtkode.SetFocus
 ElseIf txtnama.Text = "" Then
    MsgBox "NAMA JENIS TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    txtnama.SetFocus
Else
If Cek = True Then
    SQL = "Select nama from tbjenis where nama='" & txtnama.Text & "'"
       Set recordset = koneksi.Execute(SQL, , adCmdText)
       If Not recordset.EOF Then
            MsgBox "DATA JENIS SUDAH ADA, SILAHKAN CEK KEMBALI" + Chr(13) + "NOTE:", 64, "Konfirmasi"
            txtnama.SetFocus
        Else
            Query = "call TambahJenis('" & txtkode & "','" & txtnama & "','" & Form_utama.StatusBar1.Panels(1).Text & "',now())"
            koneksi.Execute Query, , adCmdText
            MsgBox "DATA JENIS BERHASIL DISIMPAN" + Chr(13) + "NOTE:", 64, "Konfirmasi"
            Call Form_Activate
            Me.FGmerk.Refresh
        End If
    Else
    SQL = "Select nama from tbjenis where nama='" & txtnama.Text & "'"
       Set recordset = koneksi.Execute(SQL, , adCmdText)
       If Not recordset.EOF Then
            MsgBox "DATA JENIS SUDAH ADA, SILAHKAN CEK KEMBALI" + Chr(13) + "NOTE:", 64, "Konfirmasi"
            txtnama.SetFocus
        Else
            Query = "call EditJenis('" & txtkode & "','" & txtnama & "','" & Form_utama.StatusBar1.Panels(1).Text & "',now())"
            koneksi.Execute Query, , adCmdText
            MsgBox "DATA JENIS BERHASIL DIUBAH" + Chr(13) + "NOTE:", 64, "Konfirmasi"
            Call Form_Activate
            Me.FGmerk.Refresh
       End If
    End If
End If
End Sub

Private Sub cmdTAMBAH_Click()
    Cek = True
    Aktifkan
    BERSIHKAN
    KodeOto
    txtnama.SetFocus
End Sub

Private Sub cmdUBAH_Click()
    Aktifkan
    Cek = False
End Sub
Private Sub FGmerk_DblClick()
    Dim barisGrid As String
    barisGrid = FGmerk.Row
    
    If FGmerk.Rows <> 1 Then
        txtkode.Text = _
            FGmerk.TextMatrix(barisGrid, 1)
        txtnama.Text = _
            FGmerk.TextMatrix(barisGrid, 2)
    Else
        Exit Sub
    End If
    cmdUBAH.Enabled = True
    cmdHAPUS.Enabled = True
    cmdBATAL.Enabled = True
    cmdKELUAR.Enabled = False
End Sub

Private Sub Form_Activate()
   TampilGrid
   tdkAktifkan
   BERSIHKAN
End Sub

Private Sub tdkAktifkan()
    txtnama.Locked = True
    cmdTAMBAH.Enabled = True
    cmdSIMPAN.Enabled = False
    cmdHAPUS.Enabled = False
    cmdUBAH.Enabled = False
    cmdKELUAR.Enabled = True
    cmdBATAL.Enabled = False
End Sub
Private Sub BERSIHKAN()
    txtkode.Text = ""
    txtnama.Text = ""
End Sub
Private Sub Aktifkan()
    txtnama.Locked = False
    cmdTAMBAH.Enabled = False
    cmdSIMPAN.Enabled = True
    cmdKELUAR.Enabled = False
    cmdBATAL.Enabled = True
End Sub
Public Sub RecTerakhir()
On Error Resume Next
    Query = "select max(idjenis)from tbjenis"
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
           KODEMERK = Val(Right(txtkode, 5) + 1)
            If KODEMERK >= 0 And KODEMERK <= 9 Then
                Me.txtkode.Text = "J-" + "0000" & Trim(Str(KODEMERK))
            ElseIf KODEMERK >= 10 And KODEMERK <= 99 Then
                Me.txtkode.Text = "J-" + "000" & Trim(Str(KODEMERK))
            ElseIf KODEMERK >= 100 And KODEMERK <= 999 Then
                Me.txtkode.Text = "J-" + "00" & Trim(Str(KODEMERK))
            End If
        Else
           Me.txtkode.Text = "J-" + "00001"
        End If
End Sub
Sub TampilGrid()
    Dim BARIS As Integer
    
    FGmerk.Clear
     Call AktifGridMerk
     
         
    FGmerk.Rows = 2
     BARIS = 0
     
     
   Set rs_JENIS = New ADODB.recordset
   Query = "select *from tbJENIS WHERE nama LIKE '%" & txtCARI.Text & "%'"
   Set rs_JENIS = koneksi.Execute(Query, , adCmdText)
   
     If rs_JENIS.EOF Then

     Else
         With rs_JENIS
            .MoveFirst
         Do While Not .EOF
            
            BARIS = BARIS + 1
            FGmerk.Rows = BARIS + 1
            FGmerk.TextMatrix(BARIS, 0) = BARIS
            FGmerk.TextMatrix(BARIS, 1) = .Fields("idJenis")
            FGmerk.TextMatrix(BARIS, 2) = .Fields("nama")
         .MoveNext
         Loop
         End With
     End If
End Sub
Sub AktifGridMerk()
    With FGmerk
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
        .Text = "KODE JENIS"
        .CellFontBold = True
        .ColWidth(1) = 1600
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 2
        .Row = 0
        .Text = "NAMA JENIS"
        .CellFontBold = True
        .ColWidth(2) = 5000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
    End With
End Sub

Private Sub Form_Click()
    pop1.Visible = False
End Sub

Private Sub Form_Load()
TampilGrid
End Sub

Private Sub Frame1_Click()
    pop1.Visible = False
End Sub

Private Sub txtcari_Change()
Query = "select * from tbjenis WHERE nama LIKE '%" & txtCARI.Text & "%' ORDER BY idJENIS"
     Set recordset = koneksi.Execute(Query, , adCmdText)
        If recordset.EOF Then
            MsgBox "TIDAK MENEMUKAN KODE JENIS! " _
            & " - " & txtCARI.Text & " - dalam tabel", _
            vbInformation, "Informasi"
            
            txtCARI.Text = ""
            txtCARI.SetFocus
        Else
          Call TampilGrid
        End If
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer)
Call BlokKarakter(KeyAscii)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtnama_Change()
    pop1.Visible = True
    txtnama.SelStart = Len(txtnama.Text)
    namaJenis
End Sub

Private Sub txtnama_KeyPress(KeyAscii As Integer)
Call BlokKarakter(KeyAscii)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Public Function nvl(isi, kondisi)
    If IsNull(isi) = True Then
        nvl = kondisi
    Else
        nvl = isi
    End If
End Function
Private Sub namaJenis()
   Query = "select * from tbjenis where nama like '%" & txtnama.Text & "%'"
        Set rs_BARANG = koneksi.Execute(Query, , adCmdText)
        If rs_BARANG.EOF Then
            lvnm1.ListItems.Clear
        Else
          rs_BARANG.MoveFirst
                        lvnm1.ListItems.Clear
                        Do While Not rs_BARANG.EOF
                            Set Item = lvnm1.ListItems.Add(, , rs_BARANG.Fields("nama"))
                            rs_BARANG.MoveNext
                        Loop
        End If
End Sub
