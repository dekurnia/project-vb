VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Lap_Terima 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Data Barang"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6900
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
   Icon            =   "LihatTerima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport crtgl 
      Left            =   6720
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tanggal Penerimaan Barang"
      Height          =   2415
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      Begin VB.ComboBox Combo1 
         Height          =   390
         Left            =   1920
         TabIndex        =   7
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&LIHAT"
         Height          =   855
         Left            =   4800
         Picture         =   "LihatTerima.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   480
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Konf"
         Height          =   270
         Left            =   480
         TabIndex        =   6
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sampai"
         Height          =   240
         Left            =   480
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dari"
         Height          =   240
         Left            =   480
         TabIndex        =   3
         Top             =   480
         Width           =   510
      End
   End
End
Attribute VB_Name = "Lap_Terima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If Combo1.Text = "" Then
MsgBox "PILIH STATUS!", vbInformation + vbOKOnly, _
                "Informasi"
ElseIf Combo1.Text = "ALL" Then
     Dim SQL2 As String
        SQL2 = ""
        SQL2 = "select tbtmpterima1.kdkirim from tbtmpterima1,tbtmpterima,tbbarang,tbmerk,tbdistributor where tbtmpterima1.kdkirim=tbtmpterima.kdkirim and tbtmpterima.kdbarang=tbbarang.kdbarang and tbtmpterima1.kddistributor=tbdistributor.kddistributor and tbbarang.idmerk=tbmerk.idmerk " _
                & " AND tbtmpterima1.tgl_terima >= '" & Format$(DTPicker1.Value, "yyyy-mm-dd") & "'  AND tbtmpterima1.tgl_terima <= '" & Format$(DTPicker2.Value, "yyyy-mm-dd") & "'  order by tbtmpterima1.kdkirim"
                        
            Set rs_BARANG = koneksi.Execute(SQL2)
            If rs_BARANG.BOF Then
                MsgBox "DATA TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                "Informasi"
                Command1.SetFocus
            Else
            
                With Me.crtgl
                    .ReportFileName = App.Path & "\Report\TerimaBrg.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    .SelectionFormula = "{TBtmpterima1.TGL_TERIMA}>=#" & DTPicker1.Value & "#" _
                                & " and {TBtmpterima1.TGL_TERIMA}<=#" & DTPicker2.Value & "# "
                    .Action = 1
                End With
            End If

ElseIf Combo1.Text = "Y" Then
    
        SQL2 = ""
         SQL2 = "select tbtmpterima1.kdkirim from tbtmpterima1,tbtmpterima,tbbarang,tbmerk,tbdistributor where tbtmpterima1.kdkirim=tbtmpterima.kdkirim and tbtmpterima.kdbarang=tbbarang.kdbarang and tbtmpterima1.kddistributor=tbdistributor.kddistributor and tbbarang.idmerk=tbmerk.idmerk " _
                & " AND tbtmpterima1.tgl_terima >= '" & Format$(DTPicker1.Value, "yyyy-mm-dd") & "'  AND tbtmpterima1.tgl_terima <= '" & Format$(DTPicker2.Value, "yyyy-mm-dd") & "' and tbtmpterima1.flag='Y' order by tbtmpterima1.kdkirim"
                       
            Set rs_BARANG = koneksi.Execute(SQL2)
            If rs_BARANG.BOF Then
                MsgBox "DATA TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                "Informasi"
                Command1.SetFocus
            Else
            
                With Me.crtgl
                    .ReportFileName = App.Path & "\Report\TerimaBrg.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    .SelectionFormula = "{TBtmpterima1.TGL_TERIMA}>=#" & DTPicker1.Value & "#" _
                                & " and {TBtmpterima1.TGL_TERIMA}<=#" & DTPicker2.Value & "# " _
                                & "And {tbtmpterima1.flag}='" & Combo1.Text & "'"
                    .Action = 1
                End With
            End If
ElseIf Combo1.Text = "T" Then
   
        SQL2 = ""
         SQL2 = "select tbtmpterima1.kdkirim from tbtmpterima1,tbtmpterima,tbbarang,tbmerk,tbdistributor where tbtmpterima1.kdkirim=tbtmpterima.kdkirim and tbtmpterima.kdbarang=tbbarang.kdbarang and tbtmpterima1.kddistributor=tbdistributor.kddistributor and tbbarang.idmerk=tbmerk.idmerk " _
                & " AND tbtmpterima1.tgl_terima >= '" & Format$(DTPicker1.Value, "yyyy-mm-dd") & "'  AND tbtmpterima1.tgl_terima <= '" & Format$(DTPicker2.Value, "yyyy-mm-dd") & "' and tbtmpterima1.flag='N' order by tbtmpterima1.kdkirim"
                       
            Set rs_BARANG = koneksi.Execute(SQL2)
            If rs_BARANG.BOF Then
                MsgBox "DATA TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                "Informasi"
                Command1.SetFocus
            Else
            
                With Me.crtgl
                    .ReportFileName = App.Path & "\Report\TerimaBrg.rpt"
                    .WindowState = crptMaximized
                    .RetrieveDataFiles
                    .SelectionFormula = "{TBtmpterima1.TGL_TERIMA}>=#" & DTPicker1.Value & "#" _
                                & " and {TBtmpterima1.TGL_TERIMA}<=#" & DTPicker2.Value & "# " _
                                & "And {tbtmpterima1.flag}='" & Combo1.Text & "'"
                    .Action = 1
                End With
            End If
End If
End Sub

Private Sub Form_Load()
        Combo1.AddItem "ALL"
        Combo1.AddItem "Y"
        Combo1.AddItem "N"
End Sub
