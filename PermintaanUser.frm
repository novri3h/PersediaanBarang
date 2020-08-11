VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form PermintaanUser 
   Caption         =   "Permintaan Barang Dari Customer"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9555
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
   ScaleHeight     =   5910
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CR 
      Left            =   5520
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox NomorReffUser 
      Height          =   350
      Left            =   6240
      TabIndex        =   20
      Top             =   1320
      Width           =   3015
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6105
      Left            =   9720
      TabIndex        =   18
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton CmdBantuan 
      Caption         =   "Lihat &Kode Barang"
      Height          =   375
      Left            =   3360
      TabIndex        =   17
      Top             =   5400
      Width           =   2000
   End
   Begin VB.TextBox PersonCus 
      Height          =   350
      Left            =   1440
      TabIndex        =   13
      Top             =   1320
      Width           =   3000
   End
   Begin VB.TextBox NamaCus 
      Height          =   350
      Left            =   6240
      TabIndex        =   12
      Top             =   960
      Width           =   3000
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   3060
   End
   Begin MSDataGridLib.DataGrid DG 
      Bindings        =   "PermintaanUser.frx":0000
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   9255
      _ExtentX        =   16325
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
      Left            =   7200
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
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   1000
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   5400
      Width           =   1000
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   5400
      Width           =   1000
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   9240
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nomor Reff User"
      Height          =   225
      Left            =   4920
      TabIndex        =   19
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label TotalKrm 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7080
      TabIndex        =   16
      Top             =   5400
      Width           =   645
   End
   Begin VB.Label TanggalMnt 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   350
      Left            =   3000
      TabIndex        =   15
      Top             =   240
      Width           =   1250
   End
   Begin VB.Label NomorMnt 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   840
      TabIndex        =   14
      Top             =   240
      Width           =   1245
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   9240
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Contact Person"
      Height          =   225
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   1185
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Nama Customer"
      Height          =   225
      Left            =   4920
      TabIndex        =   10
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Kode Customer"
      Height          =   225
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label LblKet 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7800
      TabIndex        =   8
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nomor"
      Height          =   225
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal"
      Height          =   225
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Width           =   615
   End
   Begin VB.Label TotalMnt 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   5400
      Width           =   645
   End
End
Attribute VB_Name = "PermintaanUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Call Koneksi
ADO.ConnectionString = PathData
ADO.RecordSource = "TMPMintaUser"
ADO.Refresh
Set DG.DataSource = ADO
DG.Refresh
Call AutoMnt
TanggalMnt = Date
Call TabelKosong
ADO.Recordset.MoveFirst
CmdSimpan.Enabled = False
End Sub

Private Sub Form_Load()
Call Koneksi
'tampilkan kode dan nama customer dalam combo
RSCustomer.Open "Customer", Conn
Combo1.Clear
Do While Not RSCustomer.EOF
    Combo1.AddItem RSCustomer!KodeCus & Space(7) & RSCustomer!NamaCus
    RSCustomer.MoveNext
Loop
Call KondisiAwal
End Sub

'.menampilkan nomor permintaan secara otomatis
'berdasarkan tanggal

Private Sub AutoMnt()
Call Koneksi
RSPermintaanUser.Open "select * from PermintaanUser Where NomorMnt In(Select Max(NomorMnt)From PermintaanUser)Order By NomorMnt Desc", Conn
RSPermintaanUser.Requery
Dim Urutan As String * 10
Dim Hitung As Long
With RSPermintaanUser
    If .EOF Then
        Urutan = "MT" + Format(Date, "yymmdd") + "01"
        NomorMnt = Urutan
    Else
        If Mid(!NomorMnt, 3, 6) <> Format(Date, "yymmdd") Then
            Urutan = "MT" + Format(Date, "yymmdd") + "01"
        Else
            Hitung = Right(!NomorMnt, 2) + 1
            Urutan = "MT" + Format(Date, "yymmdd") + Right("00" & Hitung, 2)
        End If
    End If
    NomorMnt = Urutan
End With
End Sub

'mencari identitas customor berdasarkan
'3 digit pertama dalam combo
Private Sub COMBO1_Click()
Call Koneksi
RSCustomer.Open "select * from Customer where kodeCus='" & Left(Combo1, 3) & "'", Conn
If Not RSCustomer.EOF Then
    NamaCus = RSCustomer!NamaCus
    PersonCus = RSCustomer!PersonCus
    NomorReffUser.Enabled = True
Else
    MsgBox "Kode Customer tidak terdaftar"
    Combo1.SetFocus
End If

End Sub

'mencari identitas customer dapat juga diketik dalam combo
Private Sub COMBO1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Combo1 = "" Then
        MsgBox "Kode Customer wajib diisi"
        Combo1.SetFocus
        Exit Sub
    ElseIf Combo1 <> "" Then
        Call Koneksi
        RSCustomer.Open "select * from Customer where kodeCus='" & Left(Combo1, 3) & "'", Conn
        If Not RSCustomer.EOF Then
            COMBO1_Click
            NomorReffUser.Enabled = True
            NomorReffUser.SetFocus
        Else
            MsgBox "Kode Customer tidak terdaftar"
            NamaCus = ""
            PersonCus = ""
            Combo1.SetFocus
            Combo1 = ""
            Exit Sub
        End If
    End If
End If

End Sub

'grid hanya dapat diisi angka
Private Sub dg_Keypress(Keyascii As Integer)
If DG.Col = 1 Then
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack Or Keyascii = vbKeyReturn) Then Keyascii = 0
ElseIf DG.Col = 4 Then
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack Or Keyascii = vbKeyReturn) Then Keyascii = 0
End If
End Sub

Private Sub CmdBatal_Click()
Combo1 = ""
NamaCus = ""
PersonCus = ""
NomorReffUser = ""
TotalMnt = ""
TotalKrm = ""
LblStok = ""
LblKet = ""
Call TabelKosong
End Sub

Private Sub CmdTutup_Click()
Unload Me
End Sub

'transaksi dalam grid
Private Sub DG_AfterColEdit(ByVal ColIndex As Integer)
    If DG.Col = 1 Then
        If Len(ADO.Recordset!Kode) < 3 Then
            MsgBox "Kode Harus 3 digit" & Chr(13) & _
            "contoh 001,002 dan seterusnya"
            DG.Col = 1
            Exit Sub
        End If
    
        Call Koneksi
        RSBarang.Open "Select * from Barang where KodeBrg='" & ADO.Recordset!Kode & "'", Conn
        'cari data barang yang kodenya di ketik dalam grid
        If Not RSBarang.EOF Then
            ADO.Recordset!Kode = RSBarang!kodebrg
            ADO.Recordset!Nama = RSBarang!namabrg
            ADO.Recordset!stok = RSBarang!jumlahbrg
            DG.Col = 4
            DG.Refresh
            Exit Sub
        End If
    End If
    
    'indikasi terpenhi atau tidaknya permintaan
    If DG.Col = 4 Then
        ADO.Recordset!qtymnt = ADO.Recordset!qtymnt
        
        If ADO.Recordset!qtymnt > ADO.Recordset!stok Then
            ADO.Recordset!dikirim = ADO.Recordset!stok
            ADO.Recordset!ket = "Stok Kurang" & Space(2) & ADO.Recordset!qtymnt - ADO.Recordset!stok
        
        ElseIf ADO.Recordset!qtymnt = ADO.Recordset!stok Then
            ADO.Recordset!dikirim = ADO.Recordset!qtymnt
            ADO.Recordset!ket = "Terpenuhi"
        
        ElseIf ADO.Recordset!qtymnt < ADO.Recordset!stok Then
            ADO.Recordset!dikirim = ADO.Recordset!qtymnt
            ADO.Recordset!ket = "Terpenuhi"
        
        End If
        
        ADO.Recordset.Update
        ADO.Recordset.MoveNext
        DG.Col = 1
        DG.Refresh
        Call CariTotalMnt
        If TotalMnt <> "" Then
            CmdSimpan.Enabled = True
            CmdBatal.Enabled = True
            CmdTutup.Enabled = True
        End If
        Call CariTotalKrm
    End If
End Sub

'menampilkan indikasi ketersediaan barang
'secara keseluruhan
Sub Keterangan()
Call Koneksi
Dim ket As New ADODB.Recordset
ket.Open "select count(ket) as ketemu from TMPMintaUser where ket like '%Stok Kurang%'", Conn
ket.Requery

If ket!ketemu > 0 Then
    LblKet = "Stok Kurang"
Else
    LblKet = "Terpenuhi"
End If

End Sub

'mencari jumlah total dalam grid
Function CariTotalKrm()
ADO.Recordset.MoveFirst
TTL = 0
Do While Not ADO.Recordset.EOF And ADO.Recordset!dikirim <> vbNullString
    TTL = TTL + ADO.Recordset!dikirim
    ADO.Recordset.MoveNext
    If TTL = 0 Then
        TotalKrm = 0
    Else
        TotalKrm = Format(TTL, "#,###")
    End If
Loop
End Function

Function CariTotalMnt()
ADO.Recordset.MoveFirst
TTL = 0
Do While Not ADO.Recordset.EOF And ADO.Recordset!qtymnt <> vbNullString
    TTL = TTL + ADO.Recordset!qtymnt
    ADO.Recordset.MoveNext
    TotalMnt = Format(TTL, "#,###")
Loop
End Function


Private Sub CmdSimpan_Click()
Call Keterangan

    If Combo1 = "" Or NomorReffUser = "" Then
        MsgBox "data belum lengkap"
        Exit Sub
    End If

Pesan = MsgBox("Data sudah benar..?", vbYesNo)
If Pesan = vbYes Then

    Call Koneksi
    
    'simpan ke tabel PermintaanUser
    Dim Simpan As String
    Simpan = "insert into PermintaanUser(nomorMnt,tanggalMnt,kodecus,nomorreffuser,totalMnt,TotalKrm,kodepmk,ket,KetKirim) values " & _
    "('" & NomorMnt & "','" & TanggalMnt & "','" & Left(Combo1, 3) & "','" & NomorReffUser & "','" & TotalMnt & "','" & TotalKrm & "','" & Menu.STBar.Panels(1).Text & "','" & LblKet & "','Belum Dikirim')"
    Conn.Execute Simpan
        
    'simpan ke tabel DetailMintaUser
    ADO.Recordset.MoveFirst
    Do While Not ADO.Recordset.EOF And ADO.Recordset!Kode <> vbNullString
        Dim simpan2 As String
        simpan2 = "insert into DetailMintaUser(nomorMnt,KODEBRG,stok,QTYMnt,dikirim,ket) values " & _
        "('" & NomorMnt & "','" & ADO.Recordset!Kode & "','" & ADO.Recordset!stok & "','" & ADO.Recordset!qtymnt & "','" & ADO.Recordset!dikirim & "','" & ADO.Recordset!ket & "')"
        Conn.Execute simpan2
    ADO.Recordset.MoveNext
    Loop
    Form_Activate
    Call Kosongkan
    Call TabelKosong
'    Call CetakPermintaanUser
End If
End Sub

Sub Kosongkan()
Call KosongkanCus
TotalMnt = ""
TotalKrm = ""
LblKet = ""
End Sub

Sub CetakPermintaanUser()
    CR.ReportFileName = App.Path & "\master minta user.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

'kosongkan tabel temporer sebelum digunakan dalam transaksi
Function TabelKosong()
On Error Resume Next
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

Private Sub NomorReffUser_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If NomorReffUser = "" Then
        NomorReffUser = "Kosong"
        DG.SetFocus
        DG.Col = 1
    Else
        DG.SetFocus
        DG.Col = 1
    End If
End If

End Sub

Sub BukaCus()
Combo1.Enabled = True
NamaCus.Enabled = False
PersonCus.Enabled = False
NomorReffUser.Enabled = True
End Sub

Sub TutupCus()
Combo1.Enabled = False
NamaCus.Enabled = False
PersonCus.Enabled = False
'NomorReffUser.Enabled = False
End Sub

Sub KondisiAwal()
Call TutupCus
Call KosongkanCus
Call Kosongkan
Combo1.Enabled = True
End Sub

Sub KosongkanCus()
Combo1 = ""
NamaCus = ""
PersonCus = ""
NomorReffUser = ""
End Sub

'=====================================

Private Sub CmdBantuan_Click()
If CmdBantuan.Caption = "Lihat &Kode Barang" Then
    Me.Width = 12800
    Call Tengah
    CmdBantuan.Caption = "Tutup &Kode Barang"
    Call Koneksi
    RSBarang.Open "select * from barang order by namabrg", Conn
    List1.Clear
    Do While Not RSBarang.EOF
        List1.AddItem RSBarang!kodebrg & vbTab & RSBarang!jumlahbrg & vbTab & RSBarang!namabrg
        RSBarang.MoveNext
    Loop
Else
    Me.Width = 9800
    Call Tengah
    CmdBantuan.Caption = "Lihat &Kode Barang"
End If
End Sub

Public Sub Tengah()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
End Sub



