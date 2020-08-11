VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Penerimaan 
   Caption         =   "Penerimaan Barang Dari Supplier"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6420
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNomorBon 
      Height          =   350
      Left            =   4800
      TabIndex        =   16
      Top             =   600
      Width           =   1500
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   5280
      Width           =   1000
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   5280
      Width           =   1000
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5280
      Width           =   1000
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1680
      TabIndex        =   5
      Top             =   600
      Width           =   1500
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   330
      Left            =   3360
      Top             =   5280
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DG 
      Bindings        =   "Penerimaan.frx":0000
      Height          =   3375
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Nomor"
         Caption         =   "Nomor"
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
         DataField       =   "Kode"
         Caption         =   "Kode"
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
      BeginProperty Column02 
         DataField       =   "Nama"
         Caption         =   "Nama"
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
      BeginProperty Column03 
         DataField       =   "StokAwal"
         Caption         =   "StokAwal"
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
      BeginProperty Column04 
         DataField       =   "Diterima"
         Caption         =   "Diterima"
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
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   2505,26
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nomor Dari Supplier"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3240
      TabIndex        =   15
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label LblTotal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label PersonSpl 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   9
      Top             =   1320
      Width           =   4605
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Contact Person"
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label NamaSpl 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   7
      Top             =   960
      Width           =   4605
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama Supplier"
      Height          =   345
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode Supplier"
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label TanggalTrm 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label NomorTrm 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tanggal"
      Height          =   345
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nomor"
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "Penerimaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Call Koneksi
ADO.ConnectionString = PathData
ADO.RecordSource = "TMPTerima"
ADO.Refresh
Set DG.DataSource = ADO
DG.Refresh
Call Auto
TanggalTrm = Date
Call TabelKosong
ADO.Recordset.MoveFirst
End Sub

'saat form di load..
'tampilkan kode supplier dalam combo
Private Sub Form_Load()
Call Koneksi
RSSupplier.Open "supplier", Conn
Combo1.Clear
Do While Not RSSupplier.EOF
    Combo1.AddItem RSSupplier!Kodespl
    RSSupplier.MoveNext
Loop
End Sub

'menampilkan nomor penerimaan otomatis
'berdasarkan tanggal
Private Sub Auto()
Call Koneksi
RSPenerimaan.Open "select * from Penerimaan Where NomorTrm In(Select Max(NomorTrm)From Penerimaan)Order By NomorTrm Desc", Conn
RSPenerimaan.Requery
Dim Urutan As String * 10
Dim Hitung As Long
With RSPenerimaan
    If .EOF Then
        Urutan = "TR" + Format(Date, "yymmdd") + "01"
        NomorTrm = Urutan
    Else
        If Mid(!NomorTrm, 3, 6) <> Format(Date, "yymmdd") Then
            Urutan = "TR" + Format(Date, "yymmdd") + "01"
        Else
            Hitung = Right(!NomorTrm, 2) + 1
            Urutan = "TR" + Format(Date, "yymmdd") + Right("00" & Hitung, 2)
        End If
    End If
    NomorTrm = Urutan
End With
End Sub

'dalam grid hanya dapat diisi angka
Private Sub dg_Keypress(Keyascii As Integer)
If DG.Col = 1 Then
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack Or Keyascii = vbKeyReturn) Then Keyascii = 0
ElseIf DG.Col = 4 Then
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack Or Keyascii = vbKeyReturn) Then Keyascii = 0
End If
End Sub


Private Sub CmdBatal_Click()
Combo1 = ""
TxtNomorBon = ""
LblTotal = ""
NamaSpl = ""
PersonSpl = ""
Call TabelKosong
Combo1.SetFocus
End Sub

Private Sub CmdTutup_Click()
Unload Me
End Sub

'menampilkan identitas supplier
'saat combo di klik
Private Sub COMBO1_Click()
Call Koneksi
RSSupplier.Open "select * from Supplier where kodespl='" & Combo1 & "'", Conn
If Not RSSupplier.EOF Then
    NamaSpl = RSSupplier!NamaSpl
    PersonSpl = RSSupplier!PersonSpl
Else
    MsgBox "Kode Supplier tidak terdaftar"
    Combo1.SetFocus
End If
End Sub

'kode supplier dalam combo dapat dipilih
'dan dapat diketik lalu menekan enter
Private Sub COMBO1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Combo1 = "" Then
        MsgBox "Kode supplier wajib diisi"
        Combo1.SetFocus
        Exit Sub
    ElseIf Combo1 <> "" Then
        Call Koneksi
        RSSupplier.Open "select * from supplier where kodespl='" & Combo1 & "'", Conn
        If Not RSSupplier.EOF Then
            COMBO1_Click
            TxtNomorBon.SetFocus
        Else
            MsgBox "Kode supplier tidak terdaftar"
            NamaSpl = ""
            PersonSpl = ""
            Combo1.SetFocus
            Combo1 = ""
            Exit Sub
        End If
    End If
End If
End Sub

Private Sub Command2_Click()
Total = ""
Combo1 = ""
NamaDpt = ""
PersonDpt = ""
Call TabelKosong
Combo1.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

'transaksi dalam grid
Private Sub DG_AfterColEdit(ByVal ColIndex As Integer)
    If DG.Col = 1 Then
        If Len(ADO.Recordset!Kode) < 3 Then
            MsgBox "Kode Harus 3 digit"
            DG.Col = 1
            Exit Sub
        End If
    
        Call Koneksi
        RSBarang.Open "Select * from Barang where KodeBrg='" & ADO.Recordset!Kode & "'", Conn
        'menampilkan data barang jika kodenya ditemukan
        If Not RSBarang.EOF Then
            ADO.Recordset!Kode = RSBarang!kodebrg
            ADO.Recordset!Nama = RSBarang!namabrg
            ADO.Recordset!stokawal = RSBarang!jumlahbrg
            DG.Col = 4
            DG.Refresh
            Exit Sub
        End If
    End If
    
    If DG.Col = 4 Then
        ADO.Recordset!diterima = ADO.Recordset!diterima
        ADO.Recordset.Update
        ADO.Recordset.MoveNext
        DG.Col = 1
        Call TotalBarang
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

'mencari total barang dalam grid
Function TotalBarang()
ADO.Recordset.MoveFirst
TTL = 0
Do While Not ADO.Recordset.EOF And ADO.Recordset!diterima <> 0
    TTL = TTL + ADO.Recordset!diterima
    ADO.Recordset.MoveNext
    LblTotal = Format(TTL, "#,###")
Loop
End Function


Private Sub CmdSimpan_Click()
If Combo1 = "" Or TxtNomorBon = "" Or LblTotal = "" Then
    MsgBox "data belum lengkap"
    If Combo1 = "" Then
        Combo1.SetFocus
    ElseIf TxtNomorBon = "" Then
        TxtNomorBon.SetFocus
    End If
    Exit Sub
End If

Call Koneksi
Dim simpan1 As String
'simpan transaksi dalam grid ke tabel penerimaan (hanya satu record)
simpan1 = "insert into Penerimaan(nomorTrm,tanggalTrm,kodespl,nomorbon,totaltrm,kodepmk) values " & _
"('" & NomorTrm & "','" & TanggalTrm & "','" & Combo1 & "','" & TxtNomorBon & "','" & LblTotal & "','" & Menu.STBar.Panels(1).Text & "')"
Conn.Execute simpan1

'simpan transaksi dalam grid ke tabel detailterima (beberapa record / berulang)
ADO.Recordset.MoveFirst
Do While Not ADO.Recordset.EOF And ADO.Recordset!Kode <> vbNullString
    Dim simpan2 As String
    simpan2 = "insert into DETAILterima(nomorTrm,KODEBRG,Stokawal,QTYTrm,stokakhir) values " & _
    "('" & NomorTrm & "','" & ADO.Recordset!Kode & "','" & ADO.Recordset!stokawal & "','" & ADO.Recordset!diterima & "','" & ADO.Recordset!stokawal + ADO.Recordset!diterima & "')"
    Conn.Execute simpan2
ADO.Recordset.MoveNext
Loop

'tambah data stok barang yang kodenya diketik dalam grid
ADO.Recordset.MoveFirst
Do While Not ADO.Recordset.EOF
    If ADO.Recordset!Kode <> vbNullString Then
        Call Koneksi
        RSBarang.Open "Select * from Barang where Kodebrg='" & ADO.Recordset!Kode & "'", Conn
        If Not RSBarang.EOF Then
            'tambah barang jika kodenya ditemukan
            Dim TambahBarang1 As String
            TambahBarang1 = "update barang set jumlahbrg='" & RSBarang!jumlahbrg + ADO.Recordset!diterima & "' where kodebrg='" & ADO.Recordset!Kode & "'"
            Conn.Execute (TambahBarang1)
        End If
    End If
ADO.Recordset.MoveNext
Loop
    
Form_Activate
Call Kosongkan
Call TabelKosong
Combo1.SetFocus
End Sub

Sub Kosongkan()
Combo1 = ""
NamaSpl = ""
PersonSpl = ""
TxtNomorBon = ""
LblTotal = ""
End Sub

'cetak laporan penerimaan
Sub CetakPenerimaan()
    CR.ReportFileName = App.Path & "\Penerimaan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

'hapus isi transaksi dalam grid sebelum digunakan
Function TabelKosong()
If ADO.Recordset.RecordCount <> 0 Then
    ADO.Recordset.MoveFirst
    Do While Not ADO.Recordset.EOF
        ADO.Recordset.Delete
        ADO.Recordset.MoveNext
    Loop
    For i = 1 To 10
        ADO.Recordset.AddNew
        ADO.Recordset!nomor = i
        ADO.Recordset.Update
    Next i
    ADO.Recordset.MoveFirst
    DG.Col = 1
End If
End Function

Private Sub Label5_Click()

End Sub

'nomor BON dari supplier
Private Sub TxtNomorBon_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If TxtNomorBon = "" Then
        TxtNomorBon = "Kosong"
        DG.SetFocus
        DG.Col = 1
    End If
End If
End Sub
