VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Lap_returUn 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Data Barang"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8295
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
   Icon            =   "Lap_returUn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&ALL"
      Height          =   2295
      Left            =   6480
      Picture         =   "Lap_returUn.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   6255
      Begin VB.CommandButton Command1 
         Caption         =   "&LIHAT"
         Height          =   855
         Left            =   4800
         Picture         =   "Lap_returUn.frx":5E8A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
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
         Format          =   86704131
         CurrentDate     =   37461
         MinDate         =   40909
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1920
         TabIndex        =   7
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
         Format          =   86704131
         CurrentDate     =   37461
         MinDate         =   40909
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sampai"
         Height          =   240
         Left            =   480
         TabIndex        =   9
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dari"
         Height          =   240
         Left            =   480
         TabIndex        =   8
         Top             =   960
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cetak No Bukti"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton cmdlihat 
         Caption         =   "&LIHAT"
         Height          =   855
         Left            =   6120
         Picture         =   "Lap_returUn.frx":6532
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbterima 
         Height          =   390
         Left            =   2280
         TabIndex        =   1
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No Retur Unit"
         Height          =   270
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   1380
      End
   End
   Begin Crystal.CrystalReport crRtUTg 
      Left            =   720
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crRtUKo 
      Left            =   120
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrRtUA 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "Lap_returUn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdlihat_Click()
    Dim SQL1 As String
        SQL1 = ""
        SQL1 = "select kdreturUN from tbreturUN where konfirm='Y' and kdreturUN='" & cmbterima & "' order by kdreturUN"
            Set rs_BARANG = koneksi.Execute(SQL1)
            If rs_BARANG.BOF Then
                MsgBox "DATA TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                "Informasi"
                cmbterima.SetFocus
            Else
                With Me.crRtUTg
                    .ReportFileName = App.Path & "\Report\FORM RETUR UNIT.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    .SelectionFormula = "{tbreturUN.kdreturUN}='" & cmbterima.Text & "'"
                    .Action = 1
                End With
            End If
End Sub

Private Sub Command1_Click()
             Dim SQL2 As String
        SQL2 = ""
        SQL2 = "select tbreturUN.*,tbUNIT.namaUNIT from tbreturUN,tbdetreturUN,tbUNIT,tbkirimUn where tbreturUN.kdkirimUN=tbkirimun.kdkirimUN and tbreturUN.kdreturUN=tbdetreturUN.kdreturUN and tbkirimUN.kdunit=tbunit.kdunit" _
                & " AND tbreturUN.tglreturUN >= '" & Format$(DTPicker1.Value, "yyyy-mm-dd") & "'  AND tbreturUN.tglreturUN <= '" & Format$(DTPicker2.Value, "yyyy-mm-dd") & "'  order by tbreturUN.kdreturUN"
                        
            Set rs_BARANG = koneksi.Execute(SQL2)
            If rs_BARANG.BOF Then
                MsgBox "DATA TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                "Informasi"
                Command1.SetFocus
            Else
                With Me.crRtUKo
                    .ReportFileName = App.Path & "\Report\LAPORAN DATA RETUR UNIT.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    .SelectionFormula = "{tbreturUN.tglreturUN}>=#" & DTPicker1.Value & "#" _
                                & " and {tbreturUN.tglreturUN}<=#" & DTPicker2.Value & "# "
                    .Action = 1
                End With
            End If
End Sub

Private Sub Command2_Click()
            With Me.CrRtUA
                    .ReportFileName = App.Path & "\Report\LAPORAN DATA RETUR UNIT.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    .Action = 1
    End With
End Sub

Private Sub Form_Load()
             '---Aktifkan Table Merk untuk Combo merk
     cmbterima.Clear
    Set rs_TERIMA = New ADODB.recordset
    rs_TERIMA.Open "select kdReturun from tbreturun where konfirm='Y' order by kdreturun", koneksi, adOpenDynamic, adLockOptimistic
    Do Until rs_TERIMA.EOF
       cmbterima.AddItem rs_TERIMA("kdreturun")
       rs_TERIMA.MoveNext
    Loop
End Sub
