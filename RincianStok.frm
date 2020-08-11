VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RincianStok 
   Caption         =   "Laporan Stok Barang Per Tanggal"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9240
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
   ScaleHeight     =   4980
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSemua 
      Caption         =   "Cetak"
      Enabled         =   0   'False
      Height          =   300
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   1500
   End
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "Cetak"
      Enabled         =   0   'False
      Height          =   300
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   1500
   End
   Begin VB.CommandButton CmdTerima 
      Caption         =   "Cetak"
      Enabled         =   0   'False
      Height          =   300
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1500
   End
   Begin VB.ListBox List3 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   1500
   End
   Begin VB.ListBox List2 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1500
   End
   Begin VB.ListBox List1 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1500
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   350
      Left            =   1800
      TabIndex        =   2
      Top             =   4440
      Width           =   1000
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   2880
      TabIndex        =   0
      Top             =   4440
      Width           =   1000
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "RincianStok.frx":0000
      Height          =   4215
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   7
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
         Caption         =   "Nama Barang"
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
         DataField       =   "Terima"
         Caption         =   "Terima"
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
      BeginProperty Column05 
         DataField       =   "Keluar"
         Caption         =   "Keluar"
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
      BeginProperty Column06 
         DataField       =   "StokAkhir"
         Caption         =   "StokAkhir"
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
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2505,26
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   615,118
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   794,835
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5880
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
   Begin Crystal.CrystalReport CR 
      Left            =   5280
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label6 
      Caption         =   "Tgl Terima dan Keluar"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Tgl Keluar"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Tgl Terima"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "RincianStok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Call Koneksi
Adodc1.ConnectionString = PathData
Adodc1.RecordSource = "TMPStokTgl"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
Call Auto
TanggalMnt = Date
End Sub

Private Sub Form_Load()
Call Koneksi
Dim stokterima As New ADODB.Recordset
stokterima.Open "select distinct tanggaltrm from penerimaan", Conn
List1.Clear
Do While Not stokterima.EOF
    List1.AddItem stokterima!TanggalTrm
    stokterima.MoveNext
Loop

Dim stokKeluar As New ADODB.Recordset
stokKeluar.Open "select distinct tanggalklr from pengeluaran", Conn
List2.Clear
Do While Not stokKeluar.EOF
    List2.AddItem stokKeluar!TanggalKlr
    stokKeluar.MoveNext
Loop

Dim stokAkhir As New ADODB.Recordset
stokAkhir.Open "select distinct tanggaltrm from penerimaan,pengeluaran where penerimaan.tanggaltrm=pengeluaran.tanggalklr", Conn
List3.Clear
Do While Not stokAkhir.EOF
    List3.AddItem stokAkhir!TanggalTrm
    stokAkhir.MoveNext
Loop

CmdBatal_Click
End Sub

Private Sub Auto()
Call Koneksi
RSPermintaanUser.Open "select * from permintaanuser Where NomorMnt In(Select Max(NomorMnt)From permintaanuser)Order By NomorMnt Desc", Conn
RSPermintaanUser.Requery
Dim Urutan As String * 10
Dim Hitung As Long
With RSPermintaanUser
    If .EOF Then
        Urutan = Format(Date, "yymmdd") + "0001"
        NomorMnt = Urutan
    Else
        If Left(!NomorMnt, 6) <> Format(Date, "yymmdd") Then
            Urutan = Format(Date, "yymmdd") + "0001"
        Else
            Hitung = (!NomorMnt) + 1
            Urutan = Format(Date, "yymmdd") + Right("0000" & Hitung, 4)
        End If
    End If
    NomorMnt = Urutan
End With
End Sub


Private Sub CmdBatal_Click()
Combo1 = ""
Call TabelKosong
TxtTotal = ""
End Sub

Private Sub CmdTutup_Click()
Unload Me
End Sub

Sub TabelKosong()
Call Koneksi
Dim hapus As String
hapus = "delete * from TMPStokTgl"
Conn.Execute hapus
Form_Activate
End Sub


Private Sub CmdTampilkan_Click()
If Combo1 = "" Then
    MsgBox "pilih jumlah barang minimal dalam combo"
    Combo1.SetFocus
    Exit Sub
Else
    Call TabelKosong
    RSBarang.Open "select * from barang where val(jumlahbrg)<=" & Val(Combo1) & "", Conn
    RSBarang.Requery
    If RSBarang.EOF Then
        MsgBox "data tidak ditemukan"
        Call TabelKosong
    Else
        RSBarang.MoveFirst
        nomor = 0
        Do While Not RSBarang.EOF
            nomor = nomor + 1
            Adodc1.Recordset.AddNew
            Adodc1.Recordset!nomor = nomor
            Adodc1.Recordset!Kode = RSBarang!kodebrg
            Adodc1.Recordset!Nama = RSBarang!namabrg
            Adodc1.Recordset!JUMLAH = RSBarang!jumlahbrg
            Adodc1.Recordset.Update
            RSBarang.MoveNext
        Loop
        'Call TotalItem
        TxtTotal.Enabled = False
    End If
End If
End Sub

Private Sub List1_Click()
    CmdKeluar.Enabled = False
    CmdSemua.Enabled = False
    CmdTerima.Enabled = True
    Call TabelKosong
    Dim tampilkan As New ADODB.Recordset
    tampilkan.Open "select distinct barang.kodebrg,namabrg,stokawal,qtytrm,'0' as keluar,stokakhir from barang,detailterima,penerimaan where barang.kodebrg=detailterima.kodebrg and penerimaan.nomortrm=detailterima.nomortrm and cdate(penerimaan.tanggaltrm)='" & List1 & "'", Conn
    tampilkan.Requery
    If tampilkan.EOF Then
        MsgBox "data tidak ditemukan"
        Call TabelKosong
    Else
        tampilkan.MoveFirst
        nomor = 0
        Do While Not tampilkan.EOF
            nomor = nomor + 1
            Adodc1.Recordset.AddNew
            Adodc1.Recordset!nomor = nomor
            Adodc1.Recordset!Kode = tampilkan!kodebrg
            Adodc1.Recordset!Nama = tampilkan!namabrg
            Adodc1.Recordset!stokawal = tampilkan!stokawal
            Adodc1.Recordset!terima = tampilkan!qtytrm
            Adodc1.Recordset!keluar = tampilkan!keluar
            Adodc1.Recordset!stokAkhir = tampilkan!stokAkhir
            Adodc1.Recordset.Update
            tampilkan.MoveNext
        Loop
    End If
End Sub

Private Sub List2_Click()
    CmdTerima.Enabled = False
    CmdSemua.Enabled = False
    CmdKeluar.Enabled = True
   Call TabelKosong
    Dim tampilkan As New ADODB.Recordset
    tampilkan.Open "select barang.kodebrg,namabrg,stokawal,'0' as Terima, dikirim,stokakhir from barang,detailkeluar,pengeluaran where barang.kodebrg=detailkeluar.kodebrg and pengeluaran.nomorklr=detailkeluar.nomorklr and cdate(pengeluaran.tanggalklr)='" & List2 & "'", Conn
    tampilkan.Requery
    If tampilkan.EOF Then
        MsgBox "data tidak ditemukan"
        Call TabelKosong
    Else
        tampilkan.MoveFirst
        nomor = 0
        Do While Not tampilkan.EOF
            nomor = nomor + 1
            Adodc1.Recordset.AddNew
            Adodc1.Recordset!nomor = nomor
            Adodc1.Recordset!Kode = tampilkan!kodebrg
            Adodc1.Recordset!Nama = tampilkan!namabrg
            Adodc1.Recordset!stokawal = tampilkan!stokawal
            Adodc1.Recordset!terima = tampilkan!terima
            Adodc1.Recordset!keluar = tampilkan!dikirim
            Adodc1.Recordset!stokAkhir = tampilkan!stokAkhir
            
            Adodc1.Recordset.Update
            tampilkan.MoveNext
        Loop
    End If
End Sub

Private Sub List3_Click()
    CmdTerima.Enabled = False
    CmdKeluar.Enabled = False
    CmdSemua.Enabled = True
    Call TabelKosong
    Dim tampilkan As New ADODB.Recordset
    tampilkan.Open "select distinct barang.kodebrg,namabrg,(barang.jumlahbrg+detailterima.stokakhir)-detailkeluar.stokakhir as stokawal,detailterima.qtytrm,detailkeluar.dikirim as keluar,(barang.jumlahbrg+detailterima.stokakhir)-detailkeluar.stokakhir + detailterima.qtytrm -detailkeluar.dikirim as stokakhir from barang,detailterima,detailkeluar,penerimaan,pengeluaran where barang.kodebrg=detailterima.kodebrg and barang.kodebrg=detailkeluar.kodebrg and penerimaan.tanggaltrm=pengeluaran.tanggalklr and cdate(tanggaltrm)='" & List3 & "'", Conn
    tampilkan.Requery
    If tampilkan.EOF Then
        MsgBox "data tidak ditemukan"
        Call TabelKosong
    Else
        tampilkan.MoveFirst
        nomor = 0
        Do While Not tampilkan.EOF
            nomor = nomor + 1
            Adodc1.Recordset.AddNew
            Adodc1.Recordset!nomor = nomor
            Adodc1.Recordset!Kode = tampilkan!kodebrg
            Adodc1.Recordset!Nama = tampilkan!namabrg
            Adodc1.Recordset!stokawal = tampilkan!stokawal
            Adodc1.Recordset!terima = tampilkan!qtytrm
            Adodc1.Recordset!keluar = tampilkan!keluar
            Adodc1.Recordset!stokAkhir = tampilkan!stokAkhir
            Adodc1.Recordset.Update
            tampilkan.MoveNext
        Loop
    End If
End Sub

Private Sub CmdKeluar_Click()
    CR.ReportFileName = App.Path & "\lap stok per tanggal.rpt"
    CR.Formulas(0) = "ambilTANGGAL='" & List2 & "'"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1

End Sub

Private Sub CmdSemua_Click()
    CR.ReportFileName = App.Path & "\lap stok per tanggal.rpt"
    CR.Formulas(0) = "ambilTANGGAL='" & List3 & "'"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub CmdTerima_Click()
    CR.ReportFileName = App.Path & "\lap stok per tanggal.rpt"
    CR.Formulas(0) = "ambilTANGGAL='" & List1 & "'"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

