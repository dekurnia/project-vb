Attribute VB_Name = "mdl"
Option Explicit
Public koneksi As ADODB.Connection
Public recordset As ADODB.recordset
Public recordsett As ADODB.recordset
Public rsTBMERK As ADODB.recordset
Public rsTBJENIS As ADODB.recordset
Public rsTBBM As ADODB.recordset
Public rs_JENIS As ADODB.recordset
Public rs_BARANG As ADODB.recordset
Public rs_STOK As ADODB.recordset
Public rs_DIS As ADODB.recordset
Public rs_TERIMA As ADODB.recordset
Public rs_TERIMA1 As ADODB.recordset
Public rs_KIRIM As ADODB.recordset
Public rs_KIRIM1 As ADODB.recordset
Public rs_retur As ADODB.recordset
Public rs_Hapus As ADODB.recordset
Public rs_unit As ADODB.recordset
Public rs_user As ADODB.recordset
Public Query, SQL, SQL2, SQL3, sql4, sql5, sqleditki, Pesan, NoBukti As String
Public Cek As Boolean

Public Sub BukaDatabase()
Set koneksi = New ADODB.Connection
 koneksi.ConnectionString = "provider=msdasql.1;" & _
   "persist security info=false; data source = invent_db;" & _
   "initial catalog=invent_db"
   koneksi.Open "invent_db"
End Sub
Public Sub BlokKarakter(KeyAscii) ' Mencegah Penulisan Karakter `!'$%^&()_+=';<>?,/\[]{}|:
    
    Dim X As String
    X = "`!'$%^&()_+=';<>?/\[]{}|:'*""@#~-"
    If InStr(1, X, Chr(KeyAscii)) > 0 Then
    KeyAscii = 0
    Beep
    End If
End Sub
Public Sub HanyaNomor(KeyAscii)
    
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 44 Or KeyAscii = 13 Or KeyAscii = vbKeyBack) Then
    KeyAscii = 0
    Beep
    End If
End Sub

