VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form PengeluaranBarang 
   Caption         =   "Pengeluaran Barang"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CR 
      Left            =   3840
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox NomorReffUser 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6120
      TabIndex        =   21
      Top             =   1440
      Width           =   2500
   End
   Begin VB.TextBox KodeCus 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1680
      TabIndex        =   20
      Top             =   1080
      Width           =   2500
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "Tutup"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   5520
      Width           =   800
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   5520
      Width           =   800
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1500
   End
   Begin VB.TextBox NamaCus 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6120
      TabIndex        =   5
      Top             =   1080
      Width           =   2500
   End
   Begin VB.TextBox PersonCus 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1680
      TabIndex        =   4
      Top             =   1440
      Width           =   2500
   End
   Begin MSDataGridLib.DataGrid DG 
      Bindings        =   "PengeluaranBarang.frx":0000
      Height          =   3255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5741
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
         DataField       =   "Stok"
         Caption         =   "Stok"
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
         DataField       =   "QtyMnt"
         Caption         =   "QtyMnt"
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
         DataField       =   "Dikirim"
         Caption         =   "Dikirim"
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
         DataField       =   "Ket"
         Caption         =   "Ket"
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
            Object.Visible         =   0   'False
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   3495,118
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   375
      Left            =   6600
      Top             =   240
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
   Begin VB.Line Line2 
      X1              =   120
      X2              =   8640
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nomor Reff User"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4560
      TabIndex        =   22
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label TanggalKlr 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5160
      TabIndex        =   19
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal Pengeluaran"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3360
      TabIndex        =   18
      Top             =   600
      Width           =   1650
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Nomor Pengeluaran"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label TotalMnt 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   16
      Top             =   5520
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal Permintaan"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3360
      TabIndex        =   15
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nomor Permintaan"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label LblKet 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   13
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Kode Customer"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Nama Customer"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4560
      TabIndex        =   11
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Contact Person"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1185
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8640
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label TanggalMnt 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5160
      TabIndex        =   9
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label TotalKrm 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Top             =   5520
      Width           =   645
   End
   Begin VB.Label NomorKlr 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   7
      Top             =   480
      Width           =   1500
   End
End
Attribute VB_Name = "PengeluaranBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Call Koneksi
ADO.ConnectionString = PathData
ADO.RecordSource = "TMPKeluarUser"
ADO.Refresh
Set DG.DataSource = ADO
DG.Refresh
TanggalKlr = Date
'tampilkan data permintaan user yang belum dikirim
RSPermintaanUser.Open "select * from permintaanuser where ketkirim='Belum Dikirim'", Conn
Combo1.Clear
Do While Not RSPermintaanUser.EOF
    Combo1.AddItem RSPermintaanUser!NomorMnt
    RSPermintaanUser.MoveNext
Loop
Call TabelKosong
End Sub

Private Sub Form_Load()
Call KondisiAwal
Call TabelKosong
End Sub

Sub TabelKosong()
Call Koneksi
Dim hapus As String
hapus = "delete * from TMPKELUARuSER"
Conn.Execute hapus
End Sub

Private Sub CmdSimpan_Click()
If Combo1 = "" Then
    MsgBox "Pilih nomor permintaan di combo1"
    Combo1.SetFocus
    Exit Sub
Else
    Call Koneksi
    RSPermintaanUser.Open "select * from permintaanuser where nomormnt='" & Combo1 & "'", Conn
    If Not RSPermintaanUser.EOF Then
        Dim edit As String
        'edit data permintaan bahwa nomor ini SUDAH DIKIRIM
        edit = "update permintaanuser set ketkirim='Sudah Dikirim' where nomormnt='" & Combo1 & "'"
        Conn.Execute edit

        
        Dim Simpan As String
        'simpan ke tabel pengeluaran
        Simpan = "insert into pengeluaran(nomorklr,tanggalklr,kodecus,nomorbon,totalmnt,TotalKrm,kodepmk,ket,KetKirim) values " & _
        "('" & NomorKlr & "','" & TanggalKlr & "','" & KodeCus & "','" & NomorReffUser & "','" & TotalMnt & "','" & TotalKrm & "','" & Menu.STBar.Panels(1).Text & "','" & LblKet & "','Sudah Dikirim')"
        Conn.Execute Simpan

        'simpan ke tabel detailkeluar
        ADO.Recordset.MoveFirst
        Do While Not ADO.Recordset.EOF
            Dim simpan2 As String
            simpan2 = "insert into Detailkeluar(nomorklr,KODEBRG,stokawal,QTYMnt,dikirim,Stokakhir,ket) values " & _
            "('" & NomorKlr & "','" & ADO.Recordset!Kode & "','" & ADO.Recordset!stok & "','" & ADO.Recordset!qtymnt & "','" & ADO.Recordset!dikirim & "','" & ADO.Recordset!stok - ADO.Recordset!dikirim & "','" & ADO.Recordset!ket & "')"
            Conn.Execute simpan2
        ADO.Recordset.MoveNext
        Loop
        
        'kurangi jumlah barang
        ADO.Recordset.MoveFirst
        Do While Not ADO.Recordset.EOF
            If ADO.Recordset!Kode <> vbNullString Then
                Call Koneksi
                RSBarang.Open "Select * from Barang where Kodebrg='" & ADO.Recordset!Kode & "'", Conn
                If Not RSBarang.EOF Then
                    Dim KurangiStokBarang As String
                    KurangiStokBarang = "update barang set jumlahbrg='" & RSBarang!jumlahbrg - ADO.Recordset!dikirim & "' where kodebrg='" & ADO.Recordset!Kode & "'"
                    Conn.Execute (KurangiStokBarang)
                End If
            End If
        ADO.Recordset.MoveNext
        Loop
        
        Form_Activate
        Call KondisiAwal
        'Call CetakPengeluaranBarang
    End If
End If
End Sub

Sub CetakPengeluaranBarang()
    CR.ReportFileName = App.Path & "\master pengeluaran.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

'nomor pengeluaran akan secara otomatis
'diambil dari nomor permintaan
'hanya dibedakan 2 huruf depannya saja
Private Sub COMBO1_Click()
NomorKlr = "KL" + Right(Combo1, 8)
Call Koneksi
Dim RSCari As New ADODB.Recordset
'mencari dan menampilkan data permintaan
RSCari.Open "select * from permintaanuser where nomormnt='" & Combo1 & "'", Conn
If Not RSCari.EOF Then
    TanggalMnt = RSCari!TanggalMnt
    NomorReffUser = RSCari!NomorReffUser
    TotalMnt = RSCari!TotalMnt
    TotalKrm = RSCari!TotalKrm
    LblKet = RSCari!ket
    'mencari dan menampilkan data customer
    RSCustomer.Open "select * from customer where kodecus='" & RSCari!KodeCus & "'", Conn
    If Not RSCustomer.EOF Then
        KodeCus = RSCari!KodeCus
        NamaCus = RSCustomer!NamaCus
        PersonCus = RSCustomer!PersonCus
    End If
End If
'jika data ditemukan, tampilkan datanya dalam grid
ADO.ConnectionString = PathData
ADO.RecordSource = "SELECT BARANG.KODEBRG AS KODE,NAMABRG AS NAMA,STOK,QTYMNT,DIKIRIM,KET FROM BARANG,DETAILMINTAUSER WHERE BARANG.KODEBRG=DETAILMINTAUSER.KODEBRG AND NOMORMNT='" & Combo1 & "'"
ADO.Refresh
Set DG.DataSource = ADO
DG.Refresh
End Sub

Private Sub CmdBatal_Click()
Call KondisiAwal
Form_Activate
Combo1.SetFocus
End Sub

Private Sub CmdTutup_Click()
Unload Me
End Sub



Sub TutupCus()
KodeCus.Enabled = False
NamaCus.Enabled = False
PersonCus.Enabled = False
NomorReffUser.Enabled = False
End Sub

Sub KondisiAwal()
TanggalKlr = Date
Call TutupCus
Call KosongkanCus
TanggalMnt = ""
TanggalKlr = ""
NomorReffUser = ""
TotalMnt = ""
TotalKrm = ""
LblKet = ""
NomorKlr = ""
Combo1 = ""
End Sub

Sub KosongkanCus()
KodeCus = ""
NamaCus = ""
PersonCus = ""
End Sub


