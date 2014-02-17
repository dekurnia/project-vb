VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Lap_ReturSup 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Data Barang"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8340
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
   Icon            =   "Lap_ReturSup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Cetak No Bukti"
      Height          =   1335
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   7935
      Begin VB.ComboBox cmbterima 
         Height          =   390
         Left            =   2280
         TabIndex        =   9
         Top             =   480
         Width           =   3495
      End
      Begin VB.CommandButton cmdlihat 
         Caption         =   "&LIHAT"
         Height          =   855
         Left            =   6120
         Picture         =   "Lap_ReturSup.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No Retur Dist"
         Height          =   270
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   1395
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   6255
      Begin VB.CommandButton Command1 
         Caption         =   "&LIHAT"
         Height          =   855
         Left            =   4800
         Picture         =   "Lap_ReturSup.frx":5E8A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   960
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   83099651
         CurrentDate     =   37461
         MinDate         =   40909
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   1440
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   83099651
         CurrentDate     =   37461
         MinDate         =   40909
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dari"
         Height          =   240
         Left            =   480
         TabIndex        =   6
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sampai"
         Height          =   240
         Left            =   480
         TabIndex        =   5
         Top             =   1440
         Width           =   735
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&ALL"
      Height          =   2295
      Left            =   6600
      Picture         =   "Lap_ReturSup.frx":6532
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   1575
   End
   Begin Crystal.CrystalReport crRetSTG 
      Left            =   840
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crturSKRe 
      Left            =   240
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrReturSA 
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "Lap_ReturSup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdlihat_Click()
        Dim SQL1 As String
        SQL1 = ""
        SQL1 = "select kdretursup from tbretursup where konfirm='Y' and kdretursup='" & cmbterima & "' order by kdretursup"
            Set rs_BARANG = koneksi.Execute(SQL1)
            If rs_BARANG.BOF Then
                MsgBox "DATA TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                "Informasi"
                cmbterima.SetFocus
            Else
                With Me.crturSKRe
                    .ReportFileName = App.Path & "\Report\tbreturs.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    .SelectionFormula = "{tbretursup.kdretursup}='" & cmbterima.Text & "'"
                    .Action = 1
                End With
            End If
End Sub

Private Sub Command1_Click()
         Dim SQL2 As String
        SQL2 = ""
        SQL2 = "select tbretursup.*,tbdistributor.namadistributor from tbretursup,tbdetretursup,tbdistributor,tbterima where tbretursup.kdkirim=tbterima.kdkirim and tbretursup.kdretursup=tbdetretursup.kdretursup and tbterima.kddistributor=tbdistributor.kddistributor" _
                & " AND tbretursup.tglretur >= '" & Format$(DTPicker1.Value, "yyyy-mm-dd") & "'  AND tbretursup.tglretur <= '" & Format$(DTPicker2.Value, "yyyy-mm-dd") & "'  order by tbretursup.kdretursup"
                        
            Set rs_BARANG = koneksi.Execute(SQL2)
            If rs_BARANG.BOF Then
                MsgBox "DATA TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                "Informasi"
                Command1.SetFocus
            Else
                With Me.crRetSTG
                    .ReportFileName = App.Path & "\Report\TBReturAll2.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    .SelectionFormula = "{tbretursup.tglretur}>=#" & DTPicker1.Value & "#" _
                                & " and {tbretursup.tglretur}<=#" & DTPicker2.Value & "# "
                    .Action = 1
                End With
            End If
End Sub

Private Sub Command2_Click()
        With Me.CrReturSA
                    .ReportFileName = App.Path & "\Report\TBReturAll2.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    .Action = 1
    End With
End Sub

Private Sub Form_Load()
         '---Aktifkan Table Merk untuk Combo merk
     cmbterima.Clear
    Set rs_TERIMA = New ADODB.recordset
    rs_TERIMA.Open "select kdReturSup from tbretursup where konfirm='Y' order by kdretursup", koneksi, adOpenDynamic, adLockOptimistic
    Do Until rs_TERIMA.EOF
       cmbterima.AddItem rs_TERIMA("kdretursup")
       rs_TERIMA.MoveNext
    Loop
End Sub
