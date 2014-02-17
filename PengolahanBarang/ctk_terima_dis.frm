VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form ctk_terima_dis 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Data Barang"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8385
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ctk_terima_dis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton cmdlihat 
         Caption         =   "&LIHAT"
         Height          =   855
         Left            =   6120
         Picture         =   "ctk_terima_dis.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbterima 
         Height          =   390
         Left            =   2760
         TabIndex        =   1
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pilih Nama Distributor"
         Height          =   270
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   2280
      End
   End
   Begin Crystal.CrystalReport crTerimaDis 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "ctk_terima_dis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
     '---Aktifkan Table Merk untuk Combo merk
     cmbterima.Clear
    Set rs_TERIMA = New ADODB.recordset
    rs_TERIMA.Open "select namaDistributor from tbdistributor  order by namaDistributor", koneksi, adOpenDynamic, adLockOptimistic
    Do Until rs_TERIMA.EOF
       cmbterima.AddItem rs_TERIMA("namaDistributor")
       rs_TERIMA.MoveNext
    Loop
    
End Sub
Private Sub cmdlihat_Click()
    Dim SQL1 As String
        SQL1 = ""
        SQL1 = "select tbdistributor.namaDistributor from tbdistributor,tbtmpterima1  " _
        & " where  namadistributor='" & cmbterima & "' and tbtmpterima1.kddistributor=tbdistributor.kddistributor"
            Set rs_BARANG = koneksi.Execute(SQL1)
            If rs_BARANG.BOF Then
                MsgBox "DATA TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                "Informasi"
                cmbterima.SetFocus
            Else
                With Me.crTerimaDis
                    .ReportFileName = App.Path & "\Report\terimabrg.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    .SelectionFormula = "{tbdistributor.namadistributor}='" & cmbterima.Text & "'"
                    .Action = 1
                End With
            End If
End Sub


