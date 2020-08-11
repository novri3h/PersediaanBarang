VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Pengeluaran 
   Caption         =   "Pengeluaran Barang"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   1500
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   1000
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   5280
      Width           =   1000
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   5280
      Width           =   1000
   End
   Begin VB.TextBox TxtNomorBon 
      Height          =   350
      Left            =   4800
      TabIndex        =   1
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
         Name            =   "MS Sans Serif"
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
      Bindings        =   "Pengeluaran.frx":0000
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   750.047
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nomor"
      Height          =   345
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tanggal"
      Height          =   345
      Left            =   3240
      TabIndex        =   15
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label NomorKLr 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   14
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label TanggalKLr 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4800
      TabIndex        =   13
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode Departemen"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama Departemen"
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label NamaDpt 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   10
      Top             =   960
      Width           =   4605
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Contact Person"
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label PersonDpt 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   8
      Top             =   1320
      Width           =   4605
   End
   Begin VB.Label LblTotal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nomor Nota"
      Height          =   345
      Left            =   3240
      TabIndex        =   6
      Top             =   600
      Width           =   1500
   End
End
Attribute VB_Name = "Pengeluaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Call Koneksi
ADO.ConnectionString = PathData
ADO.RecordSource = "TMPKeluar"
ADO.Refresh
Set DG.DataSource = ADO
DG.Refresh
Call Auto
TanggalKLr = Date
Call TabelKosong
ADO.Recordset.MoveFirst
End Sub

Private Sub Form_Load()
Call Koneksi
RSDepartemen.Open "Departemen", Conn
Combo1.Clear
Do While Not RSDepartemen.EOF
    Combo1.AddItem RSDepartemen!KodeDpt
    RSDepartemen.MoveNext
Loop
End Sub

Private Sub dg_Keypress(Keyascii As Integer)
If DG.Col = 1 Then
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack Or Keyascii = vbKeyReturn) Then Keyascii = 0
ElseIf DG.Col = 4 Then
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack Or Keyascii = vbKeyReturn) Then Keyascii = 0
End If

End Sub


Private Sub Auto()
Call Koneksi
RSPengeluaran.Open "select * from Pengeluaran Where NomorKlr In(Select Max(NomorKlr)From Pengeluaran)Order By NomorKlr Desc", Conn
RSPengeluaran.Requery
Dim Urutan As String * 10
Dim Hitung As Long
With RSPengeluaran
    If .EOF Then
        Urutan = "KR" + Format(Date, "yymmdd") + "01"
        NomorKLr = Urutan
    Else
        If Mid(!NomorKLr, 3, 6) <> Format(Date, "yymmdd") Then
            Urutan = "KR" + Format(Date, "yymmdd") + "01"
        Else
            Hitung = Right(!NomorKLr, 2) + 1
            Urutan = "KR" + Format(Date, "yymmdd") + Right("00" & Hitung, 2)
        End If
    End If
    NomorKLr = Urutan
End With
End Sub

Private Sub CmdBatal_Click()
Combo1 = ""
TxtNomorBon = ""
LblTotal = ""
Call TabelKosong
Combo1.SetFocus
End Sub

Private Sub CmdTutup_Click()
Unload Me
End Sub


Private Sub combo1_click()
Call Koneksi
RSDepartemen.Open "select * from Departemen where kodeDpt='" & Combo1 & "'", Conn
If Not RSDepartemen.EOF Then
    NamaDpt = RSDepartemen!NamaDpt
    PersonDpt = RSDepartemen!PersonDpt
Else
    MsgBox "Kode Departemen tidak terdaftar"
    Combo1.SetFocus
End If
End Sub

Private Sub combo1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Combo1 = "" Then
        MsgBox "Kode Departemen wajib diisi"
        Combo1.SetFocus
        Exit Sub
    ElseIf Combo1 <> "" Then
        Call Koneksi
        RSDepartemen.Open "select * from Departemen where kodeDpt='" & Combo1 & "'", Conn
        If Not RSDepartemen.EOF Then
            combo1_click
            TxtNomorBon.SetFocus
        Else
            MsgBox "Kode Departemen tidak terdaftar"
            NamaDpt = ""
            PersonDpt = ""
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

Private Sub DG_AfterColEdit(ByVal ColIndex As Integer)
    If DG.Col = 1 Then
        If Len(ADO.Recordset!kode) < 3 Then
            MsgBox "Kode Harus 3 digit"
            DG.Col = 1
            Exit Sub
        End If
    
        Call Koneksi
        RSBarang.Open "Select * from Barang where KodeBrg='" & ADO.Recordset!kode & "'", Conn
        If Not RSBarang.EOF Then
            ADO.Recordset!kode = RSBarang!kodebrg
            ADO.Recordset!nama = RSBarang!namabrg
            ADO.Recordset!Stokawal = RSBarang!jumlahbrg
            DG.Col = 4
            DG.Refresh
            Exit Sub
        End If
    End If
    
    If DG.Col = 4 Then
        ADO.Recordset!Keluar = ADO.Recordset!Keluar
        ADO.Recordset.Update
        ADO.Recordset.MoveNext
        DG.Col = 1
        Call TotalBarang
    End If
End Sub

Function TotalBarang()
ADO.Recordset.MoveFirst
TTL = 0
Do While Not ADO.Recordset.EOF And ADO.Recordset!Keluar <> 0
    TTL = TTL + ADO.Recordset!Keluar
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
simpan1 = "insert into Pengeluaran(nomorKlr,tanggalKlr,kodeDpt,nomorbon,totalKlr,kodepmk) values " & _
"('" & NomorKLr & "','" & TanggalKLr & "','" & Combo1 & "','" & TxtNomorBon & "','" & LblTotal & "','" & Menu.STBar.Panels(1).Text & "')"
Conn.Execute simpan1

ADO.Recordset.MoveFirst
Do While Not ADO.Recordset.EOF And ADO.Recordset!kode <> vbNullString
    Dim simpan2 As String
    simpan2 = "insert into DETAILKeluar(nomorKlr,KODEBRG,QTYKlr) values " & _
    "('" & NomorKLr & "','" & ADO.Recordset!kode & "','" & ADO.Recordset!Keluar & "')"
    Conn.Execute simpan2
ADO.Recordset.MoveNext
Loop

ADO.Recordset.MoveFirst
Do While Not ADO.Recordset.EOF
    If ADO.Recordset!kode <> vbNullString Then
        Call Koneksi
        RSBarang.Open "Select * from Barang where Kodebrg='" & ADO.Recordset!kode & "'", Conn
        If Not RSBarang.EOF Then
            'tambah barang jika kodenya ditemukan
            Dim TambahBarang1 As String
            TambahBarang1 = "update barang set jumlahbrg='" & RSBarang!jumlahbrg - ADO.Recordset!Keluar & "' where kodebrg='" & ADO.Recordset!kode & "'"
            Conn.Execute (TambahBarang1)
'        Else
'            'input data barang jika kodenya baru
'            Dim TambahBarang2 As String
'            TambahBarang2 = "Insert Into Barang(Kodebrg,NamaBrg,HargaBeli,HargaJual,JumlahBrg)" & _
'            "values('" & ADO.Recordset!Kode & "','" & ADO.Recordset!Nama & "','" & ADO.Recordset!Harga & "','" & ADO.Recordset!Harga * 1.5 & "','" & ADO.Recordset!Jumlah & "')"
'            Conn.Execute (TambahBarang2)
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
NamaDpt = ""
PersonDpt = ""
TxtNomorBon = ""
LblTotal = ""
End Sub

Sub CetakPengeluaran()
    CR.ReportFileName = App.Path & "\Pengeluaran.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Function TabelKosong()
If ADO.Recordset.RecordCount <> 0 Then
    ADO.Recordset.MoveFirst
    Do While Not ADO.Recordset.EOF
        ADO.Recordset.Delete
        ADO.Recordset.MoveNext
    Loop
    For i = 1 To 10
        ADO.Recordset.AddNew
        ADO.Recordset!Nomor = i
        ADO.Recordset.Update
    Next i
    ADO.Recordset.MoveFirst
    DG.Col = 1
End If
End Function

Private Sub TxtNomorBon_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If TxtNomorBon = "" Then
        TxtNomorBon = "Kosong"
        DG.SetFocus
        DG.Col = 1
    End If
End If
End Sub


