VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form_distributor 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Stok Barang"
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   13500
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Distributor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   13500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Halaman Master Distributor"
      ForeColor       =   &H00000000&
      Height          =   3015
      Left            =   240
      TabIndex        =   10
      Top             =   0
      Width           =   13095
      Begin VB.TextBox txtkode 
         Height          =   390
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtnama 
         Height          =   390
         Left            =   2760
         TabIndex        =   14
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox txtalamat 
         Height          =   1335
         Left            =   2760
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   13
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox txttelp 
         Height          =   390
         Left            =   9120
         MaxLength       =   12
         TabIndex        =   12
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtkontak 
         Height          =   390
         Left            =   9120
         TabIndex        =   11
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kode Distributor"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   480
         TabIndex        =   25
         Top             =   360
         Width           =   1710
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nama Distributor"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   480
         TabIndex        =   24
         Top             =   840
         Width           =   1770
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alamat"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   480
         TabIndex        =   23
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Telpon"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7080
         TabIndex        =   22
         Top             =   480
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kontak Person"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7080
         TabIndex        =   21
         Top             =   960
         Width           =   1545
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2400
         TabIndex        =   20
         Top             =   360
         Width           =   60
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2400
         TabIndex        =   19
         Top             =   840
         Width           =   60
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2400
         TabIndex        =   18
         Top             =   1320
         Width           =   60
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   8880
         TabIndex        =   17
         Top             =   480
         Width           =   60
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   8880
         TabIndex        =   16
         Top             =   960
         Width           =   60
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   13095
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
         Picture         =   "Form_Distributor.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   360
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
         Left            =   3720
         Picture         =   "Form_Distributor.frx":5E04
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "Form_Distributor.frx":646B
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSIMPAN1 
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
         Picture         =   "Form_Distributor.frx":6A8D
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   -1680
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
         Picture         =   "Form_Distributor.frx":70AF
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "Form_Distributor.frx":7721
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "Form_Distributor.frx":7D56
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtcari 
         Height          =   390
         Left            =   9600
         TabIndex        =   2
         Top             =   600
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   555
         Left            =   12360
         Picture         =   "Form_Distributor.frx":83EB
         ToolTipText     =   "Refresh"
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Masukkan Nama"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7680
         TabIndex        =   9
         Top             =   600
         Width           =   1755
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGDIST 
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   4800
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   6800
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
Attribute VB_Name = "Form_distributor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBATAL_Click()
    tdkAktif
    bersih
End Sub

Private Sub cmdHAPUS_Click()
SQL = "CALL CekDist('" & txtkode.Text & "')"
Set rs_DIS = koneksi.Execute(SQL, , adCmdText)
If Not rs_DIS.EOF Then
 MsgBox "DATA TIDAK DAPAT DIHAPUS, SEDANG DIGUNAKAN DI TABEL LAIN " + Chr(13) + "NOTE:", 64, "Konfirmasi"
 Call Form_Activate
Else
Query = "CALL HapusDist('" & txtkode.Text & "')"
    Pesan = MsgBox("Bener Neeh Mau Dihapus !" _
            , vbQuestion + vbYesNo, "Konfirmasi")
    If Pesan = vbYes Then
       Set recordset = koneksi.Execute(Query, , adCmdText)
       Call Form_Activate
       Me.FGDIST.Refresh
    End If
End If
End Sub

Private Sub cmdKELUAR_Click()
    Unload Me
End Sub

Private Sub cmdSIMPAN_Click()
If txtkode.Text = "" Then
    MsgBox "KODE DISTRIBUTOR TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    txtkode.SetFocus
 ElseIf txtnama.Text = "" Then
    MsgBox "NAMA DISTRIBUTOR TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    txtnama.SetFocus
 ElseIf txtalamat.Text = "" Then
     MsgBox "ALAMAT TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    txtalamat.SetFocus
 ElseIf txttelp.Text = "" Then
    MsgBox "TELPON TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    txttelp.SetFocus
 ElseIf txtkontak.Text = "" Then
    MsgBox "KONTAK PERSON TIDAK BOLEH KOSONG" + Chr(13) + "NOTE:", 64, "Konfirmasi"
    txtkontak.SetFocus
 Else
    If Cek = True Then
       SQL = "Select namadistributor from tbdistributor where namadistributor='" & txtnama.Text & "'"
       Set recordset = koneksi.Execute(SQL, , adCmdText)
            If Not recordset.EOF Then
                 MsgBox "DATA DISTRIBUTOR SUDAH ADA, SILAHKAN CEK KEMBALI" + Chr(13) + "NOTE:", 64, "Konfirmasi"
                 txtnama.SetFocus
             Else
                 Query = "call TambahDist('" & txtkode & "','" & txtnama & "','" & txtalamat & "','" & txttelp & "','" & txtkontak & "','" & Form_utama.StatusBar1.Panels(1).Text & "',now())"
                 koneksi.Execute Query, , adCmdText
                 MsgBox "DATA DISTRIBUTOR BERHASIL DISIMPAN" + Chr(13) + "NOTE:", 64, "Konfirmasi"
                 Call Form_Activate
                 Me.FGDIST.Refresh
                 Cek = False
             End If
    Else
      
            Query = "call EditDist('" & txtkode & "','" & txtnama & "','" & txtalamat & "','" & txttelp & "','" & txtkontak & "','" & Form_utama.StatusBar1.Panels(1).Text & "',now())"
            koneksi.Execute Query, , adCmdText
            MsgBox "DATA DISTRIBUTOR BERHASIL DIUBAH" + Chr(13) + "NOTE:", 64, "Konfirmasi"
            Call Form_Activate
            Me.FGDIST.Refresh
        
    End If
End If
End Sub

Private Sub cmdTAMBAH_Click()
    Cek = True
    Aktifkan
    bersih
    KodeOto
    txtnama.SetFocus
    
    cmdTAMBAH.Enabled = False
    cmdSIMPAN.Enabled = True
    cmdBATAL.Enabled = True
End Sub



Private Sub cmdUBAH_Click()
    Aktifkan
    Cek = False
End Sub

Private Sub FGDIST_DblClick()
    Dim barisGrid As String
    barisGrid = FGDIST.Row
    
    If FGDIST.Rows <> 1 Then
        txtkode.Text = _
            FGDIST.TextMatrix(barisGrid, 1)
        txtnama.Text = _
            FGDIST.TextMatrix(barisGrid, 2)
        txtalamat.Text = _
            FGDIST.TextMatrix(barisGrid, 3)
        txttelp.Text = _
            FGDIST.TextMatrix(barisGrid, 4)
        txtkontak.Text = _
            FGDIST.TextMatrix(barisGrid, 5)
    Else
        Exit Sub
    End If
    cmdUBAH.Enabled = True
    cmdHAPUS.Enabled = True
    cmdBATAL.Enabled = True
    cmdKELUAR.Enabled = False
    cmdSIMPAN.Enabled = False
End Sub

Private Sub Form_Activate()
    Call tampilgrid
    Call tdkAktif
    Call bersih
End Sub
Private Sub tdkAktif()
    txtnama.Locked = True
    txtalamat.Locked = True
    txttelp.Locked = True
    txtkontak.Locked = True
    
   cmdHAPUS.Enabled = False
   cmdUBAH.Enabled = False
   cmdSIMPAN.Enabled = False
   cmdBATAL.Enabled = False
   cmdTAMBAH.Enabled = True
   cmdKELUAR.Enabled = True
End Sub
Private Sub bersih()
    txtkode.Text = ""
    txtnama.Text = ""
    txtalamat.Text = ""
    txttelp.Text = ""
    txtkontak.Text = ""
End Sub
Public Sub RecTerakhir()
On Error Resume Next
    Query = "select max(kdDistributor)from tbdistributor"
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
                Me.txtkode.Text = "DS" + "0000" & Trim(Str(KODEMERK))
            ElseIf KODEMERK >= 10 And KODEMERK <= 99 Then
                Me.txtkode.Text = "DS" + "000" & Trim(Str(KODEMERK))
            ElseIf KODEMERK >= 100 And KODEMERK <= 999 Then
                Me.txtkode.Text = "DS" + "00" & Trim(Str(KODEMERK))
            End If
        Else
           Me.txtkode.Text = "DS" + "00001"
        End If
End Sub
Private Sub Aktifkan()
    txtnama.Locked = False
    txtalamat.Locked = False
    txttelp.Locked = False
    txtkontak.Locked = False
    
   cmdHAPUS.Enabled = True
   cmdUBAH.Enabled = True
   cmdSIMPAN.Enabled = True
   cmdBATAL.Enabled = True
   cmdTAMBAH.Enabled = False
End Sub
Private Sub txtalamat_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If txtalamat.Text = "" Then
            MsgBox "ALAMAT DISTRIBUTOR TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
        Else
            txttelp.SetFocus
        End If
    End If
End Sub

Private Sub txtcari_Change()
Query = "select * from tbdistributor WHERE namaDistributor LIKE '%" & txtcari.Text & "%' ORDER BY kdDistributor"
     Set recordset = koneksi.Execute(Query, , adCmdText)
        If recordset.EOF Then
            MsgBox "TIDAK MENEMUKAN NAMA DISTRIBUTOR! " _
            & " - " & txtcari.Text & " - dalam tabel", _
            vbInformation, "Informasi"
            
            txtcari.Text = ""
            txtcari.SetFocus
        Else
          Call tampilgrid
        End If
End Sub

Private Sub txtkontak_KeyPress(KeyAscii As Integer)
Call BlokKarakter(KeyAscii)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If txtkontak.Text = "" Then
            MsgBox "KONTAK PERSON DISTRIBUTOR TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
        Else
            cmdSIMPAN.SetFocus
        End If
    End If
End Sub
Private Sub txtnama_KeyPress(KeyAscii As Integer)
Call BlokKarakter(KeyAscii)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If txtnama.Text = "" Then
            MsgBox "NAMA DISTRIBUTOR TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
        Else
            txtalamat.SetFocus
        End If
    End If
End Sub

Private Sub txttelp_KeyPress(KeyAscii As Integer)
Call HanyaNomor(KeyAscii)

    If KeyAscii = 13 Then
        If txttelp.Text = "" Then
            MsgBox "TELPON DISTRIBUTOR TIDAK BOLEH KOSONG" + Chr(13) + "ULANGI LAGI", 64, "Konfirmasi"
        Else
            txtkontak.SetFocus
        End If
    End If
End Sub
Sub tampilgrid()
    Dim BARIS As Integer
    
    FGDIST.Clear
    Call AktifGridDis
     
         
    FGDIST.Rows = 2
    BARIS = 0
     
     
   Set rs_DIS = New ADODB.recordset
   Query = "select * from tbdistributor WHERE namaDistributor LIKE '%" & txtcari.Text & "%'"
   Set rs_DIS = koneksi.Execute(Query, , adCmdText)
   
     If rs_DIS.EOF Then
         
     Else
         With rs_DIS
            .MoveFirst
         Do While Not .EOF
            BARIS = BARIS + 1
            FGDIST.Rows = BARIS + 1
            FGDIST.TextMatrix(BARIS, 0) = BARIS
            FGDIST.TextMatrix(BARIS, 1) = .Fields("kdDistributor")
            FGDIST.TextMatrix(BARIS, 2) = .Fields("namaDistributor")
            FGDIST.TextMatrix(BARIS, 3) = .Fields("alamat")
            FGDIST.TextMatrix(BARIS, 4) = .Fields("telp")
            FGDIST.TextMatrix(BARIS, 5) = .Fields("kontakPerson")
            FGDIST.TextMatrix(BARIS, 6) = .Fields("user_ubah")
            FGDIST.TextMatrix(BARIS, 7) = .Fields("tgl_ubah")
         .MoveNext
         Loop
         End With
     End If
End Sub
Sub AktifGridDis()
    With FGDIST
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
        .Text = "KODE DIST"
        .CellFontBold = True
        .ColWidth(1) = 1600
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 2
        .Row = 0
        .Text = "NAMA DISTRIBUTOR"
        .CellFontBold = True
        .ColWidth(2) = 5000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 3
        .Row = 0
        .Text = "ALAMAT"
        .CellFontBold = True
        .ColWidth(3) = 5000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 4
        .Row = 0
        .Text = "TELPON"
        .CellFontBold = True
        .ColWidth(4) = 5000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 5
        .Row = 0
        .Text = "KONTAK PERSON"
        .CellFontBold = True
        .ColWidth(5) = 5000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 6
        .Row = 0
        .Text = "USER UBAH"
        .CellFontBold = True
        .ColWidth(6) = 3000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 7
        .Row = 0
        .Text = "TGL UBAH"
        .CellFontBold = True
        .ColWidth(7) = 5000
        .AllowUserResizing = flexResizeColumns
        .CellAlignment = flexAlignCenterCenter
        
    End With
End Sub

