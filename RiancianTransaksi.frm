VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form RincianTransaksi 
   Caption         =   "Rincian TRansaksi"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10185
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
   ScaleHeight     =   4740
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   1185
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1400
   End
   Begin VB.ListBox List1 
      Height          =   1185
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1400
   End
   Begin MSDataGridLib.DataGrid DG 
      Bindings        =   "RiancianTransaksi.frx":0000
      Height          =   3135
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5530
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
            ColumnWidth     =   3404,977
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
      Left            =   7920
      Top             =   4200
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.Label NotaCus 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4800
      TabIndex        =   15
      Top             =   4080
      Width           =   1755
   End
   Begin VB.Label PersonCus 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4800
      TabIndex        =   14
      Top             =   3720
      Width           =   1755
   End
   Begin VB.Label NamaCus 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   13
      Top             =   4080
      Width           =   1750
   End
   Begin VB.Label TotalMnt 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6960
      TabIndex        =   12
      Top             =   3720
      Width           =   645
   End
   Begin VB.Label LblKet 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8400
      TabIndex        =   11
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label TotalKrm 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7680
      TabIndex        =   10
      Top             =   3720
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nomor Nota"
      Height          =   225
      Left            =   3600
      TabIndex        =   9
      Top             =   4080
      Width           =   960
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Contact Person"
      Height          =   225
      Left            =   3600
      TabIndex        =   8
      Top             =   3720
      Width           =   1185
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Nama Customer"
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal"
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label TanggalMnt 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   5
      Top             =   3720
      Width           =   1750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Pengeluaran"
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Permintaan"
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "RincianTransaksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call Koneksi
RSPermintaanUser.Open "permintaanuser", Conn
List1.Clear
Do While Not RSPermintaanUser.EOF
    List1.AddItem RSPermintaanUser!NomorMnt
    RSPermintaanUser.MoveNext
Loop

RSPengeluaran.Open "pengeluaran", Conn
List2.Clear
Do While Not RSPengeluaran.EOF
    List2.AddItem RSPengeluaran!NomorKlr
    RSPengeluaran.MoveNext
Loop
End Sub

Private Sub List1_Click()
Call Koneksi
ADO.ConnectionString = PathData
ADO.RecordSource = "select barang.kodebrg as kode,namabrg as nama,stok,qtymnt,dikirim,detailmintauser.ket  from barang,detailmintauser,permintaanuser where barang.kodebrg=detailmintauser.kodebrg and permintaanuser.nomormnt=detailmintauser.nomormnt and permintaanuser.nomormnt='" & List1.Text & "'"
ADO.Refresh
Set DG.DataSource = ADO
DG.Refresh
Dim CariData As New ADODB.Recordset
CariData.Open "select * from permintaanuser where nomormnt='" & List1 & "'", Conn
NotaCus = CariData!NomorReffUser
TanggalMnt = CariData!TanggalMnt
LblKet = CariData!ket

RSCustomer.Open "select * from customer where kodecus='" & CariData!KodeCus & "'", Conn
NamaCus = RSCustomer!NamaCus
PersonCus = RSCustomer!PersonCus

Dim CariQtyMnt As New ADODB.Recordset
CariQtyMnt.Open "select sum(QtyMnt) as KetemuQtyMnt from detailmintauser where nomormnt='" & List1 & "'", Conn
TotalMnt = CariQtyMnt!ketemuQtyMnt

Dim CariDikirim As New ADODB.Recordset
CariDikirim.Open "select sum(Dikirim) as KetemuDikirim from detailmintauser where nomormnt='" & List1 & "'", Conn
TotalKrm = CariDikirim!ketemuDikirim


End Sub

Private Sub List2_Click()
Call Koneksi
ADO.ConnectionString = PathData
ADO.RecordSource = "select barang.kodebrg as kode,namabrg as nama,stokawal as Stok,qtymnt,dikirim,detailkeluar.ket  from barang,detailkeluar,pengeluaran where barang.kodebrg=detailkeluar.kodebrg and pengeluaran.nomorklr=detailkeluar.nomorklr and pengeluaran.nomorklr='" & List2.Text & "'"
ADO.Refresh
Set DG.DataSource = ADO
DG.Refresh
Dim CariData As New ADODB.Recordset
CariData.Open "select * from pengeluaran where nomorklr='" & List2 & "'", Conn
NotaCus = CariData!nomorbon
TanggalMnt = CariData!TanggalKlr
LblKet = CariData!ket

RSCustomer.Open "select * from customer where kodecus='" & CariData!KodeCus & "'", Conn
NamaCus = RSCustomer!NamaCus
PersonCus = RSCustomer!PersonCus

Dim CariQtyMnt As New ADODB.Recordset
CariQtyMnt.Open "select sum(QtyMnt) as KetemuQtyMnt from detailkeluar where nomorklr='" & List2 & "'", Conn
TotalMnt = CariQtyMnt!ketemuQtyMnt

Dim CariDikirim As New ADODB.Recordset
CariDikirim.Open "select sum(Dikirim) as KetemuDikirim from detailkeluar where nomorklr='" & List2 & "'", Conn
TotalKrm = CariDikirim!ketemuDikirim


End Sub
