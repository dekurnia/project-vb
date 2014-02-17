VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form ctk_kir_un 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Pengolahan Barang"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "ctk_kir_un.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Laporan Pengiriman Barang"
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.ComboBox Combo1 
         Height          =   390
         Left            =   2280
         TabIndex        =   8
         Top             =   1920
         Width           =   3495
      End
      Begin VB.ComboBox cmbterima 
         Height          =   390
         Left            =   2280
         TabIndex        =   2
         Top             =   480
         Width           =   3495
      End
      Begin VB.CommandButton cmdlihat 
         BackColor       =   &H00808080&
         Caption         =   "&LIHAT"
         Height          =   1335
         Left            =   6240
         Picture         =   "ctk_kir_un.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   960
         Width           =   3495
         _ExtentX        =   6165
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
         CurrentDate     =   41659.8708680556
         MinDate         =   40909
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   1440
         Width           =   3495
         _ExtentX        =   6165
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
         CurrentDate     =   41659.879525463
         MinDate         =   40909
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Konf"
         Height          =   270
         Left            =   360
         TabIndex        =   9
         Top             =   1920
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sampai"
         Height          =   240
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dari"
         Height          =   240
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Unit"
         Height          =   270
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport crKirimun 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "ctk_kir_un"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
     '---Aktifkan Table Merk untuk Combo merk
    cmbterima.Clear
    cmbterima.Text = "SEMUA UNIT"
    Set rs_KIRIM = New ADODB.recordset
    rs_KIRIM.Open "select namaunit from tbunit  order by namaunit", koneksi, adOpenDynamic, adLockOptimistic
    Do Until rs_KIRIM.EOF
       cmbterima.AddItem rs_KIRIM("namaunit")
       rs_KIRIM.MoveNext
    Loop
    '----
    Combo1.Text = "ALL"
    Combo1.AddItem "Y"
    Combo1.AddItem "N"
End Sub

Private Sub cmdlihat_Click()
        Dim SQL1 As String
        If cmbterima.Text = "SEMUA UNIT" Then
            If Combo1.Text = "ALL" Then
                SQL1 = ""
                SQL1 = "select tbTMPkirimun1.kdkirimun from tbtmpkirimun1,tbtmpkirimun,tbbarang,tbmerk,tbunit where " _
                & " tbtmpkirimun1.kdkirimun=tbtmpkirimun.kdkirimun and tbtmpkirimun.kdbarang=tbbarang.kdbarang and tbtmpkirimun1.kdunit=tbunit.kdunit and tbbarang.idmerk=tbmerk.idmerk " _
                & " AND tbtmpkirimun1.tglkirim >= '" & Format$(DTPicker1.Value, "yyyy-mm-dd") & "'  AND tbtmpkirimun1.tglkirim <= '" & Format$(DTPicker2.Value, "yyyy-mm-dd") & "'  order by tbtmpkirimun1.kdkirimun"
                   
                    Set rs_BARANG = koneksi.Execute(SQL1)
                    If rs_BARANG.BOF Then
                        MsgBox "DATA TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                        "Informasi"
                        cmbterima.SetFocus
                    Else
                        With Me.crKirimun
                            .ReportFileName = App.Path & "\Report\kirimunit.rpt"
                            .WindowState = crptMaximized
                            .RetrieveDataFiles
                            .SelectionFormula = "{tbTMPkirimun1.tglkirim}>=#" & DTPicker1.Value & "#" _
                                & " and {tbTMPkirimun1.tglkirim}<=#" & DTPicker2.Value & "# "
                            .Action = 1
                        End With
                    End If
                ElseIf Combo1.Text = "Y" Then
                    SQL1 = ""
                    SQL1 = "select tbTMPkirimun1.kdkirimun from tbtmpkirimun1,tbtmpkirimun,tbbarang,tbmerk,tbunit where " _
                    & " tbtmpkirimun1.kdkirimun=tbtmpkirimun.kdkirimun and tbtmpkirimun.kdbarang=tbbarang.kdbarang and tbtmpkirimun1.kdunit=tbunit.kdunit and tbbarang.idmerk=tbmerk.idmerk " _
                    & " AND tbtmpkirimun1.tglkirim >= '" & Format$(DTPicker1.Value, "yyyy-mm-dd") & "'  AND tbtmpkirimun1.tglkirim <= '" & Format$(DTPicker2.Value, "yyyy-mm-dd") & "'  and tbtmpkirimun1.flag='Y' order by tbtmpkirimun1.kdkirimun"
                       
                        Set rs_BARANG = koneksi.Execute(SQL1)
                        If rs_BARANG.BOF Then
                            MsgBox "DATA TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                            "Informasi"
                            cmbterima.SetFocus
                        Else
                            With Me.crKirimun
                                .ReportFileName = App.Path & "\Report\kirimunit.rpt"
                                .WindowState = crptMaximized
                                .RetrieveDataFiles
                                .SelectionFormula = "{tbTMPkirimun1.tglkirim}>=#" & DTPicker1.Value & "#" _
                                    & " and {tbTMPkirimun1.tglkirim}<=#" & DTPicker2.Value & "# " _
                                    & " and {tbtmpkirimun1.flag}='" & Combo1.Text & "'"
                                .Action = 1
                            End With
                        End If
                ElseIf Combo1.Text = "Y" Then
                    SQL1 = ""
                    SQL1 = "select tbTMPkirimun1.kdkirimun from tbtmpkirimun1,tbtmpkirimun,tbbarang,tbmerk,tbunit where " _
                    & " tbtmpkirimun1.kdkirimun=tbtmpkirimun.kdkirimun and tbtmpkirimun.kdbarang=tbbarang.kdbarang and tbtmpkirimun1.kdunit=tbunit.kdunit and tbbarang.idmerk=tbmerk.idmerk " _
                    & " AND tbtmpkirimun1.tglkirim >= '" & Format$(DTPicker1.Value, "yyyy-mm-dd") & "'  AND tbtmpkirimun1.tglkirim <= '" & Format$(DTPicker2.Value, "yyyy-mm-dd") & "'  and tbtmpkirimun1.flag='N' order by tbtmpkirimun1.kdkirimun"
                       
                        Set rs_BARANG = koneksi.Execute(SQL1)
                        If rs_BARANG.BOF Then
                            MsgBox "DATA TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                            "Informasi"
                            cmbterima.SetFocus
                        Else
                            With Me.crKirimun
                                .ReportFileName = App.Path & "\Report\kirimunit.rpt"
                                .WindowState = crptMaximized
                                .RetrieveDataFiles
                                .SelectionFormula = "{tbTMPkirimun1.tglkirim}>=#" & DTPicker1.Value & "#" _
                                    & " and {tbTMPkirimun1.tglkirim}<=#" & DTPicker2.Value & "# " _
                                    & " and {tbtmpkirimun1.flag}='" & Combo1.Text & "'"
                                .Action = 1
                            End With
                        End If
                End If
        Else
        'unit dipilih
                If Combo1.Text = "ALL" Then
                SQL1 = ""
                SQL1 = "select tbTMPkirimun1.kdkirimun from tbtmpkirimun1,tbtmpkirimun,tbbarang,tbmerk,tbunit where " _
                & " tbtmpkirimun1.kdkirimun=tbtmpkirimun.kdkirimun and tbtmpkirimun.kdbarang=tbbarang.kdbarang and tbtmpkirimun1.kdunit=tbunit.kdunit and tbbarang.idmerk=tbmerk.idmerk " _
                & " AND tbtmpkirimun1.tglkirim >= '" & Format$(DTPicker1.Value, "yyyy-mm-dd") & "'  AND tbtmpkirimun1.tglkirim <= '" & Format$(DTPicker2.Value, "yyyy-mm-dd") & "'  and tbunit.namaunit='" & cmbterima & "' order by tbtmpkirimun1.kdkirimun"
                   
                    Set rs_BARANG = koneksi.Execute(SQL1)
                    If rs_BARANG.BOF Then
                        MsgBox "DATA TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                        "Informasi"
                        cmbterima.SetFocus
                    Else
                        With Me.crKirimun
                            .ReportFileName = App.Path & "\Report\kirimunit.rpt"
                            .WindowState = crptMaximized
                            .RetrieveDataFiles
                            .SelectionFormula = "{tbTMPkirimun1.tglkirim}>=#" & DTPicker1.Value & "#" _
                                & " and {tbTMPkirimun1.tglkirim}<=#" & DTPicker2.Value & "# " _
                                & " and {tbunit.namaunit}='" & cmbterima.Text & "'"
                            .Action = 1
                        End With
                    End If
                ElseIf Combo1.Text = "Y" Then
                    SQL1 = ""
                    SQL1 = "select tbTMPkirimun1.kdkirimun from tbtmpkirimun1,tbtmpkirimun,tbbarang,tbmerk,tbunit where " _
                    & " tbtmpkirimun1.kdkirimun=tbtmpkirimun.kdkirimun and tbtmpkirimun.kdbarang=tbbarang.kdbarang and tbtmpkirimun1.kdunit=tbunit.kdunit and tbbarang.idmerk=tbmerk.idmerk " _
                    & " AND tbtmpkirimun1.tglkirim >= '" & Format$(DTPicker1.Value, "yyyy-mm-dd") & "'  AND tbtmpkirimun1.tglkirim <= '" & Format$(DTPicker2.Value, "yyyy-mm-dd") & "'  and tbtmpkirimun1.flag='Y' and tbunit.namaunit='" & cmbterima & "'  order by tbtmpkirimun1.kdkirimun"
                       
                        Set rs_BARANG = koneksi.Execute(SQL1)
                        If rs_BARANG.BOF Then
                            MsgBox "DATA TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                            "Informasi"
                            cmbterima.SetFocus
                        Else
                            With Me.crKirimun
                                .ReportFileName = App.Path & "\Report\kirimunit.rpt"
                                .WindowState = crptMaximized
                                .RetrieveDataFiles
                                .SelectionFormula = "{tbTMPkirimun1.tglkirim}>=#" & DTPicker1.Value & "#" _
                                    & " and {tbTMPkirimun1.tglkirim}<=#" & DTPicker2.Value & "# " _
                                    & " and {tbtmpkirimun1.flag}='" & Combo1.Text & "'" _
                                    & " and {tbunit.namaunit}='" & cmbterima.Text & "'"
                                .Action = 1
                            End With
                        End If
                ElseIf Combo1.Text = "Y" Then
                    SQL1 = ""
                    SQL1 = "select tbTMPkirimun1.kdkirimun from tbtmpkirimun1,tbtmpkirimun,tbbarang,tbmerk,tbunit where " _
                    & " tbtmpkirimun1.kdkirimun=tbtmpkirimun.kdkirimun and tbtmpkirimun.kdbarang=tbbarang.kdbarang and tbtmpkirimun1.kdunit=tbunit.kdunit and tbbarang.idmerk=tbmerk.idmerk " _
                    & " AND tbtmpkirimun1.tglkirim >= '" & Format$(DTPicker1.Value, "yyyy-mm-dd") & "'  AND tbtmpkirimun1.tglkirim <= '" & Format$(DTPicker2.Value, "yyyy-mm-dd") & "'  and tbtmpkirimun1.flag='N' and tbunit.namaunit='" & cmbterima & "' order by tbtmpkirimun1.kdkirimun"
                       
                        Set rs_BARANG = koneksi.Execute(SQL1)
                        If rs_BARANG.BOF Then
                            MsgBox "DATA TIDAK TERSEDIA !", vbInformation + vbOKOnly, _
                            "Informasi"
                            cmbterima.SetFocus
                        Else
                            With Me.crKirimun
                                .ReportFileName = App.Path & "\Report\kirimunit.rpt"
                                .WindowState = crptMaximized
                                .RetrieveDataFiles
                                .SelectionFormula = "{tbTMPkirimun1.tglkirim}>=#" & DTPicker1.Value & "#" _
                                    & " and {tbTMPkirimun1.tglkirim}<=#" & DTPicker2.Value & "# " _
                                    & " and {tbtmpkirimun1.flag}='" & Combo1.Text & "'" _
                                    & " and {tbunit.namaunit}='" & cmbterima.Text & "'"
                                .Action = 1
                            End With
                        End If
                End If
        End If
End Sub


