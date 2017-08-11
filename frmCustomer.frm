VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCustomer 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database II Project"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCustomer.frx":0000
   ScaleHeight     =   8820
   ScaleWidth      =   16065
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox idbarang 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   29
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton tambah 
      Caption         =   "Tambah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13920
      TabIndex        =   28
      Top             =   6960
      Width           =   1815
   End
   Begin VB.ComboBox ukuran 
      Height          =   315
      ItemData        =   "frmCustomer.frx":0342
      Left            =   2760
      List            =   "frmCustomer.frx":034F
      TabIndex        =   27
      Top             =   4080
      Width           =   1815
   End
   Begin VB.ComboBox jenis 
      Height          =   315
      ItemData        =   "frmCustomer.frx":035D
      Left            =   2760
      List            =   "frmCustomer.frx":0367
      TabIndex        =   26
      Top             =   3600
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4935
      Left            =   7080
      TabIndex        =   25
      Top             =   1560
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8705
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11760
      TabIndex        =   24
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   23
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ubah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11760
      TabIndex        =   22
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton simpan 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   21
      Top             =   6960
      Width           =   1815
   End
   Begin VB.TextBox satuan 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   20
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox jumlah 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   19
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox harga 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   18
      Top             =   6480
      Width           =   1815
   End
   Begin VB.TextBox tanggal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   17
      Top             =   7080
      Width           =   1815
   End
   Begin VB.TextBox warna 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   12
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox namabarang 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox idpelanggan 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   2280
      Width           =   3015
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   7
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox namapelanggan 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label12 
      Caption         =   "tanggal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "satuan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "jumlah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "harga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "ukuran"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "warna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "nama barang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "id pelanggan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "jenis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "id barang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "nama pelanggan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "PELANGGAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
On Error GoTo err_CmdAdd_Click
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim num_records As Integer

  ' Get the database name.
    db_file = App.Path
    If Right$(db_file, 1) <> "\" Then db_file = db_file & _
        "\"
    db_file = db_file & "OurDB.mdb"

    ' Open a connection.
    Set conn = New ADODB.Connection
    conn.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & db_file & ";" & _
        "Persist Security Info=False"
    conn.Open
    
If txtId.Text = "" Or txtFName.Text = "" Or txtNos.Text = "" Or txtMName.Text = "" Or txtLName.Text = "" Then
    MsgBox "All fields are required!", vbExclamation, "Error"
    Exit Sub
  End If


  ' Populate the table.
   conn.Execute "INSERT INTO tblcust VALUES('" & txtId.Text & "','" & txtFName.Text & "','" & txtMName.Text & "','" & txtLName.Text & "','" & _
           txtNos.Text & "')"
           
    ' See how many records the table contains.
    Set rs = conn.Execute("SELECT COUNT (*) FROM tblcust")
    num_records = rs.Fields(0)
     MsgBox "Item Added: " & num_records & " Customer in Total", _
        vbInformation, "Done"

exit_err_CmdAdd_Click:
    Exit Sub
 
err_CmdAdd_Click:
    If Err = -2147467259 Then
        MsgBox "ID Taken. Use Another", vbCritical, "Error"
    Else
        MsgBox Err.Description, vbInformation, "Proceed"
    End If
    
End Sub

Private Sub cmdClear_Click()
txtFName.Text = ""
txtMName.Text = ""
txtLName.Text = ""
txtNos.Text = ""
cntRec
End Sub

Private Sub cmdEdit_Click()
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

  ' Get the database name.
    db_file = App.Path
    If Right$(db_file, 1) <> "\" Then db_file = db_file & _
        "\"
    db_file = db_file & "OurDB.mdb"

    ' Open a connection.
    Set conn = New ADODB.Connection
    conn.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & db_file & ";" & _
        "Persist Security Info=False"
    conn.Open
If txtId.Text = "" Or txtFName.Text = "" Or txtNos.Text = "" Or txtMName.Text = "" Or txtLName.Text = "" Then
    MsgBox "All fields are required!", vbExclamation, "Error"
    Exit Sub
  End If
  
  If txtSearch.Text = "" Then
  MsgBox "Search the Item First", vbCritical, "Error"
  Exit Sub
  End If
If MsgBox("This action will modify the selected record.  Proceed?", vbYesNo, "Update") = vbYes Then
    Set rs = conn.Execute("UPDATE tblcust SET CustId='" & txtId & "', FName='" & txtFName & "',MName='" & txtMName & "',LName='" & txtLName & "',Contact_Nos='" & txtNos & "'" & _
          "WHERE CustId=" & txtSearch.Text & "")
           MsgBox "Edited Sucessfully", _
        vbInformation, "Done"
 Else
    Cancel = True
  End If

   
End Sub
Sub kosong()
namapelanggan.Text = ""
idbarang.Text = ""
idpelanggan.Text = ""
namabarang.Text = ""
jenis.Text = ""
ukuran.Text = ""
warna.Text = ""
satuan.Text = ""
harga.Text = ""
tanggal.Text = ""
End Sub

Sub tidakaktif()
namapelanggan.Enabled = False
idbarang.Enabled = False
idpelanggan.Enabled = False
namabarang.Enabled = False
jenis.Enabled = False
ukuran.Enabled = False
warna.Enabled = False
satuan.Enabled = False
harga.Enabled = False
tanggal.Enabled = False
End Sub

Private Sub cmdView_Click()
frmCustView.Show
End Sub

Sub cntRec()
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim num_records As Integer

  ' Get the database name.
    db_file = App.Path
    If Right$(db_file, 1) <> "\" Then db_file = db_file & _
        "\"
    db_file = db_file & "OurDB.mdb"

    ' Open a connection.
    Set conn = New ADODB.Connection
    conn.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & db_file & ";" & _
        "Persist Security Info=False"
    conn.Open
    
    ' See how many records the table contains.
    Set rs = conn.Execute("SELECT MAX(CustId) FROM tblcust")
  If IsNull(rs.Fields(0)) = True Then
  txtId.Text = 10
  Else
  idbarang.Text = rs.Fields(0) + 1
  End If
End Sub

Private Sub Command4_Click()
simpandata
Command3.Enabled = True
Command3.SetFocus
End Sub

Private Sub Command5_Click()
Call aktif
Call kosong
simpan.Enable = True
hapus.Enable = False
ubah.Enable = False
barang.SetFocus
End Sub

Private Sub Form_Activate()
    Call kosong
    Call aktif
    Call koneksi
    tambah.Enabled = True
    rsfrmCustomer.Open "SELECT * FROM pelanggan", conn
    Set DataGrid1.DataSource = rsfrmCustomer.DataSource
     
If Not rsfrmCustomer.EOF Then
   
DataGrid1.Columns(0).Width = 1100
DataGrid1.Columns(2).Width = 1100
DataGrid1.Columns(3).Width = 1400
DataGrid1.Columns(4).Width = 1300
DataGrid1.Columns(5).Width = 1300
DataGrid1.Columns(6).Width = 1300

 tambah.SetFocus
End If

End Sub

Private Sub simpan_Click()
If kduser.Text = "" Or frmCustomer.Text = "" Or id_barang.Text = "" Or id_pelanggan.Text = "" Or nama_pelanggan.Text = "" Or nama_barang.Text = "" Or jenis.Text = "" Or ukuran.Text = "" Or warna.Text = "" Or satuan.Text = "" Or jumlah.Text = "" Or harga.Text = "" Or tanggal.Text = "" Then
    MsgBox "Isi data dengan lengkap", , "INFORMASI"
    CmdSimpan.Enabled = True
    Else
AUF = True
     conn.Execute "insert pelanggan (id_barang,id_pelanggan,nama_pelanggan,nama_barang,jenis,ukuran,ukuran,warna,satuan,jumlah,harga,tanggal) values ('" & id_barang.Text & "','" & id_pelanggan.Text & "','" & nama_pelanggan.Text & "','" & nama_barang.Text & "','" & jenis.Text & "','" & ukuran.Text & "','" & warna.Text & "','" & satuan.Text & "','" & jumlah.Text & "','" & harga.Text & "','" & tanggal.Text & "','" & "')"
    MsgBox "Data Sudah Tersimpan", , "SAVING...."
   
    Call Form_Activate
    cmdtambah.Enabled = True
    End If
End Sub

Private Sub tambah_Click()
Call aktif

End Sub
Private Sub Form_Load()
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim num_records As Integer

    ' Get the database name.
    db_file = App.Path
    If Right$(db_file, 1) <> "\" Then db_file = db_file & _
        "\"
    db_file = db_file & "OurDB.mdb"

    ' Open a connection.
    Set conn = New ADODB.Connection
    conn.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & db_file & ";" & _
        "Persist Security Info=False"
    conn.Open
    cntRec
exit_err_CmdAdd_Click:
    Exit Sub
 
err_CmdAdd_Click:
    If Err = -2147467259 Then
        MsgBox "Table Created!"
    Else
        MsgBox "Table Exists!"
    End If

End Sub

Private Sub cmdDel_Click()
Dim ans As String
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

  ' Get the database name.
    db_file = App.Path
    If Right$(db_file, 1) <> "\" Then db_file = db_file & _
        "\"
    db_file = db_file & "OurDB.mdb"

    ' Open a connection.
    Set conn = New ADODB.Connection
    conn.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & db_file & ";" & _
        "Persist Security Info=False"
    conn.Open


If txtSearch.Text = "" Or txtId.Text = "" Then
    MsgBox "Nothing to Delete.", vbExclamation, "Error"
  
  Else
ans = MsgBox("Do you Want to Delete This Records", vbYesNo + vbQuestion, _
"Delete")
If ans = vbYes Then MsgBox "Succesfully Deleted", vbInformation, "Done"
csrc = txtSearch.Text
    Set rs = conn.Execute("DELETE *FROM tblcust WHERE CustID=" & txtSearch.Text & "")
If ans = vbNo Then Exit Sub

    End If
    
Set rs = Nothing
End Sub


Private Sub cmdGO_Click()
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

  ' Get the database name.
    db_file = App.Path
    If Right$(db_file, 1) <> "\" Then db_file = db_file & _
        "\"
    db_file = db_file & "OurDB.mdb"

    ' Open a connection.
    Set conn = New ADODB.Connection
    conn.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & db_file & ";" & _
        "Persist Security Info=False"
    conn.Open

If txtSearch.Text = "" Then
    MsgBox "Nothing to Search", vbExclamation, "Error"
    Exit Sub
  End If
    
    Set rs = conn.Execute("SELECT *FROM tblcust WHERE CustID=" & txtSearch.Text & "")
    If rs.BOF = False Or rs.EOF = False Then
    If rs.RecordCount <> 0 Then
    MsgBox "Item Found.", vbInformation, "Result"
   txtId.Text = rs.Fields("CustID").Value
   txtFName.Text = rs.Fields("FName").Value
   txtMName.Text = rs.Fields("MName").Value
   txtLName.Text = rs.Fields("LName").Value
   txtNos.Text = rs.Fields("Contact_Nos").Value
   End If
   Else
    MsgBox "Item Not Found!", vbCritical, "Result"
    
   End If
    rs.Close
    Set rs = Nothing

End Sub

Private Sub cmdMenu_Click()
Menu.Show
Unload Me
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
ElseIf KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub txtNos_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
ElseIf KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub


Private Sub txtSearch_GotFocus()
txtSearch.Text = ""
txtSearch.ForeColor = &H0
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
ElseIf KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub aktif()
idbarang.Enabled = True
namapelanggan.Enabled = True
idpelanggan.Enabled = True
namapelanggan.Text = Format(Date, "YYYY-MM-DD")
jenis.Enabled = True
ukuran.Enabled = True
warna.Enabled = True
satuan.Enabled = True
jumlah.Enabled = True
harga.Enabled = True
tanggal.Enabled = True
End Sub

Private Sub tampil()
    idbarang.Text = rsuser!id_barang
    id_pelanggan.Text = rsuser!id_pelanggan
    namapelanggan.Text = rsuser!nama_pelanggan
    namabarang.Text = rsuser!nama_barang
    jenis.Text = rsuser!jenis
    ukuran.Text = rsuser!ukuran
    warna.Text = rsuser!warna
    satuan.Text = rsuser!satuan
    jumlah.Text = rsuser!jumlah
    harga.Text = rsuser!harga
    tanggal.Text = rsuser!tanggal
End Sub
