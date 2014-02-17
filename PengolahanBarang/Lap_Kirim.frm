VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Lap_Kirim 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Barang"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6690
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
   Icon            =   "Lap_Kirim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport crKirimAlll 
      Left            =   360
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton Command1 
         Caption         =   "&LIHAT"
         Height          =   855
         Left            =   4800
         Picture         =   "Lap_Kirim.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   960
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
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
         Format          =   99287043
         CurrentDate     =   37461
         MinDate         =   40909
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
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
         Format          =   99287043
         CurrentDate     =   37461
         MinDate         =   40909
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dari"
         Height          =   240
         Left            =   480
         TabIndex        =   5
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sampai"
         Height          =   240
         Left            =   480
         TabIndex        =   4
         Top             =   1440
         Width           =   735
      End
   End
End
Attribute VB_Name = "Lap_Kirim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
        Dim SQL2 As String
        SQL2 = ""
        SQL2 = "select tbkirimun.kdkirimun from tbkirimun,tbdetkirimun,tbbarang,tbmerk,tbunit where tbkirimun.kdkirimun=tbdetkirimun.kdkirimun and tbdetkirimun.kdbarang=tbbarang.kdbarang and tbkirimun.kdunit=tbunit.kdunit and tbbarang.idmerk=tbmerk.idmerk " _
                & " AND tbkirimun.tglkirim >= '" & Format$(DTPicker1.Value, "yyyy-mm-dd") & "'  AND tbkirimun.tglkirim <= '" & Format$(DTPicker2.Value, "yyyy-mm-dd") & "'  order by tbkirimun.kdkirimun"
                        
            Set rs_BARANG = koneksi.Execute(SQL2)
            If rs_BARANG.BOF Then
                MsgBox "DATA TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                "Informasi"
                Command1.SetFocus
            Else
                With Me.crKirimAlll
                    .ReportFileName = App.Path & "\Report\kirimperiode.rpt"
                    .WindowState = crptMaximized
                    
                    .SelectionFormula = "{tbkirimun.tglkirim}>=#" & DTPicker1.Value & "#" _
                                & " and {tbkirimun.tglkirim}<=#" & DTPicker2.Value & "# "
                    .Formulas(0) = "formula1=#" & DTPicker1.Value & "#"
                    .Formulas(1) = "formula2=#" & DTPicker2.Value & "#"
                    .RetrieveDataFiles
                    .Action = 1
                End With
            End If
End Sub

