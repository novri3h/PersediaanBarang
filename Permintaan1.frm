VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Permintaan1 
   Caption         =   "Permintaan Barang Dari User (Departemen atau Customer)"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   6105
      Left            =   9720
      TabIndex        =   32
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton CmdBantuan 
      Caption         =   "Lihat &Kode Barang"
      Height          =   375
      Left            =   3480
      TabIndex        =   31
      Top             =   6000
      Width           =   2000
   End
   Begin VB.TextBox PersonCus 
      Height          =   350
      Left            =   6480
      TabIndex        =   26
      Top             =   1680
      Width           =   3000
   End
   Begin VB.TextBox NamaCus 
      Height          =   350
      Left            =   6480
      TabIndex        =   25
      Top             =   1320
      Width           =   3000
   End
   Begin VB.TextBox PersonDpt 
      Height          =   350
      Left            =   1800
      TabIndex        =   24
      Top             =   1680
      Width           =   3000
   End
   Begin VB.TextBox NamaDpt 
      Height          =   350
      Left            =   1800
      TabIndex        =   23
      Top             =   1320
      Width           =   3000
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Customer"
      Height          =   350
      Left            =   6960
      TabIndex        =   1
      Top             =   240
      Width           =   1250
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Departemen"
      Height          =   350
      Left            =   5640
      TabIndex        =   0
      Top             =   240
      Width           =   1250
   End
   Begin VB.TextBox NotaCus 
      Height          =   350
      Left            =   6480
      TabIndex        =   17
      Top             =   2040
      Width           =   3000
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   6480
      TabIndex        =   3
      Top             =   960
      Width           =   3060
   End
   Begin MSDataGridLib.DataGrid DG 
      Bindings        =   "Permintaan1.frx":0000
      Height          =   3255
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   5741
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
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   3495.118
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   375
      Left            =   240
      Top             =   6600
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   3060
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   6000
      Width           =   1000
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   6000
      Width           =   1000
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   6000
      Width           =   1000
   End
   Begin VB.TextBox NotaDpt 
      Height          =   350
      Left            =   1800
      TabIndex        =   8
      Top             =   2040
      Width           =   3000
   End
   Begin VB.Label NomorKlr 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   8280
      TabIndex        =   30
      Top             =   240
      Width           =   1245
   End
   Begin VB.Label TotalKrm 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7200
      TabIndex        =   29
      Top             =   6000
      Width           =   645
   End
   Begin VB.Label TanggalMnt 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   350
      Left            =   3000
      TabIndex        =   28
      Top             =   240
      Width           =   1250
   End
   Begin VB.Label NomorMnt 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   840
      TabIndex        =   27
      Top             =   240
      Width           =   1245
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   9480
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Permintaan Dari :"
      Height          =   195
      Left            =   4320
      TabIndex        =   22
      Top             =   315
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Nomor Nota"
      Height          =   345
      Left            =   4920
      TabIndex        =   21
      Top             =   2040
      Width           =   1500
   End
   Begin VB.Label Label10 
      Caption         =   "Contact Person"
      Height          =   345
      Left            =   4920
      TabIndex        =   20
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Label Label7 
      Caption         =   "Nama Customer"
      Height          =   345
      Left            =   4920
      TabIndex        =   19
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label Label4 
      Caption         =   "Kode Customer"
      Height          =   345
      Left            =   4920
      TabIndex        =   18
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label LblKet 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7920
      TabIndex        =   16
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nomor"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal"
      Height          =   195
      Left            =   2160
      TabIndex        =   14
      Top             =   240
      Width           =   585
   End
   Begin VB.Label Label5 
      Caption         =   "Kode Departemen"
      Height          =   345
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label Label6 
      Caption         =   "Nama Departemen"
      Height          =   345
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label Label8 
      Caption         =   "Contact Person"
      Height          =   345
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Label TotalMnt 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   6000
      Width           =   645
   End
   Begin VB.Label Label3 
      Caption         =   "Nomor Nota"
      Height          =   345
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   1500
   End
End
Attribute VB_Name = "Permintaan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
KodeBarang.Show
End Sub

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

Private Sub Form_Activate()
Call Koneksi
ADO.ConnectionString = PathData
ADO.RecordSource = "TMPMinta1"
ADO.Refresh
Set DG.DataSource = ADO
DG.Refresh
Call AutoMnt
Call AutoKlr
TanggalMnt = Date
Call TabelKosong
ADO.Recordset.MoveFirst
CmdSimpan.Enabled = False
CmdBatal.Enabled = False
'CmdTutup.Enabled = False
End Sub

Private Sub Form_Load()
Call Koneksi
RSDepartemen.Open "Departemen", Conn
Combo1.Clear
Do While Not RSDepartemen.EOF
    Combo1.AddItem RSDepartemen!Kodedpt & Space(7) & RSDepartemen!NamaDpt
    RSDepartemen.MoveNext
Loop

RSCustomer.Open "Customer", Conn
Combo2.Clear
Do While Not RSCustomer.EOF
    Combo2.AddItem RSCustomer!KodeCus & Space(7) & RSCustomer!NamaCus
    RSCustomer.MoveNext
Loop

Call KondisiAwal
'Option1.Value = True
End Sub

Private Sub AutoMnt()
Call Koneksi
RSPermintaan1.Open "select * from Permintaan1 Where NomorMnt In(Select Max(NomorMnt)From Permintaan1)Order By NomorMnt Desc", Conn
RSPermintaan1.Requery
Dim Urutan As String * 10
Dim Hitung As Long
With RSPermintaan1
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

Private Sub AutoKlr()
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

Private Sub Combo2_Click()
Call Koneksi
RSCustomer.Open "select * from Customer where kodeCus='" & Left(Combo2, 3) & "'", Conn
If Not RSCustomer.EOF Then
    NamaCus = RSCustomer!NamaCus
    PersonCus = RSCustomer!PersonCus
Else
    MsgBox "Kode Customer tidak terdaftar"
    Combo2.SetFocus
End If

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    If Combo2 = "" Then
        MsgBox "Kode Customer wajib diisi"
        Combo2.SetFocus
        Exit Sub
    ElseIf Combo2 <> "" Then
        Call Koneksi
        RSCustomer.Open "select * from Customer where kodeCus='" & Left(Combo2, 3) & "'", Conn
        If Not RSCustomer.EOF Then
            Combo2_Click
            NotaCus.SetFocus
        Else
            MsgBox "Kode Customer tidak terdaftar"
            NamaCus = ""
            PersonCus = ""
            Combo2.SetFocus
            Combo2 = ""
            Exit Sub
        End If
    End If
End If

End Sub

Private Sub DG_KeyPress(KeyAscii As Integer)

If DG.Col = 1 Then
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
ElseIf DG.Col = 4 Then
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
End If
End Sub

Private Sub CmdBatal_Click()
Combo1 = ""
NotaDpt = ""
NamaDpt = ""
PersonDpt = ""
TotalMnt = ""
TotalKrm = ""
LblStok = ""
LblKet = ""
Call TabelKosong
'Call KondisiAwal
End Sub

Private Sub CmdTutup_Click()
Unload Me
End Sub


Private Sub combo1_click()
Call Koneksi
RSDepartemen.Open "select * from Departemen where kodeDpt='" & Left(Combo1, 3) & "'", Conn
If Not RSDepartemen.EOF Then
    NamaDpt = RSDepartemen!NamaDpt
    PersonDpt = RSDepartemen!PersonDpt
Else
    MsgBox "Kode Departemen tidak terdaftar"
    Combo1.SetFocus
End If
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    If Combo1 = "" Then
        MsgBox "Kode Departemen wajib diisi"
        Combo1.SetFocus
        Exit Sub
    ElseIf Combo1 <> "" Then
        Call Koneksi
        RSDepartemen.Open "select * from Departemen where kodeDpt='" & Left(Combo1, 3) & "'", Conn
        If Not RSDepartemen.EOF Then
            combo1_click
            NotaDpt.SetFocus
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
        If Len(ADO.Recordset!Kode) < 3 Then
            MsgBox "Kode Harus 3 digit" & Chr(13) & _
            "contoh 001,002 dan seterusnya"
            DG.Col = 1
            Exit Sub
        End If
    
        Call Koneksi
        RSBarang.Open "Select * from Barang where KodeBrg='" & ADO.Recordset!Kode & "'", Conn
        If Not RSBarang.EOF Then
            ADO.Recordset!Kode = RSBarang!kodebrg
            ADO.Recordset!Nama = RSBarang!namabrg
            ADO.Recordset!stok = RSBarang!jumlahbrg
            DG.Col = 4
            DG.Refresh
            Exit Sub
        End If
    End If
    
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

Sub Keterangan()
Call Koneksi
Dim ket As New ADODB.Recordset
ket.Open "select count(ket) as ketemu from tmpminta1 where ket like '%Stok Kurang%'", Conn
ket.Requery

If ket!ketemu > 0 Then
    LblKet = "Stok Kurang"
Else
    LblKet = "Terpenuhi"
End If

End Sub

Function CariTotalKrm()
ADO.Recordset.MoveFirst
TTL = 0
Do While Not ADO.Recordset.EOF And ADO.Recordset!dikirim <> vbNullString
    TTL = TTL + ADO.Recordset!dikirim
    ADO.Recordset.MoveNext
    TotalKrm = Format(TTL, "#,###")
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
If Option1.Value = True Then
    If Combo1 = "" Or NotaDpt = "" Then
        MsgBox "data belum lengkap"
        Exit Sub
    End If
ElseIf Option2.Value = True Then
    If Combo2 = "" Or NotaCus = "" Then
        MsgBox "data belum lengkap"
        Exit Sub
    End If
End If

Pesan = MsgBox("Data sudah benar..?", vbYesNo)
If Pesan = vbYes Then

    Call Koneksi
    If Option1.Value = True Then
    
        Dim SimpanDpt As String
        SimpanDpt = "insert into Permintaan1(nomorMnt,tanggalMnt,kodeDpt,nomorbon,totalMnt,TotalKrm,kodepmk,ket) values " & _
        "('" & NomorMnt & "','" & TanggalMnt & "','" & Left(Combo1, 3) & "','" & NotaDpt & "','" & TotalMnt & "','" & TotalKrm & "','" & Menu.STBar.Panels(1).Text & "','" & LblKet & "')"
        Conn.Execute SimpanDpt
        
        Dim KeluarDpt As String
        KeluarDpt = "insert into Pengeluaran(nomorKlr,tanggalKlr,kodeDpt,nomorbon,totalMnt,TotalKlr,kodepmk,ket) values " & _
        "('" & NomorKLr & "','" & TanggalKLr & "','" & Left(Combo1, 3) & "','" & NotaDpt & "','" & TotalMnt & "','" & TotalKrm & "','" & Menu.STBar.Panels(1).Text & "','" & LblKet & "')"
        Conn.Execute KeluarDpt
        
        
    ElseIf Option2.Value = True Then
        
        Dim SimpanCus As String
        SimpanCus = "insert into Permintaan1(nomorMnt,tanggalMnt,kodecus,nomorbon,totalMnt,TotalKrm,kodepmk,ket) values " & _
        "('" & NomorMnt & "','" & TanggalMnt & "','" & Left(Combo2, 3) & "','" & NotaCus & "','" & TotalMnt & "','" & TotalKrm & "','" & Menu.STBar.Panels(1).Text & "','" & LblKet & "')"
        Conn.Execute SimpanCus
        
        
        Dim KeluarCus As String
        KeluarCus = "insert into Pengeluaran(nomorKlr,tanggalKlr,kodecus,nomorbon,totalMnt,TotalKlr,kodepmk,ket) values " & _
        "('" & NomorKLr & "','" & TanggalMnt & "','" & Left(Combo2, 3) & "','" & NotaCus & "','" & TotalMnt & "','" & TotalKrm & "','" & Menu.STBar.Panels(1).Text & "','" & LblKet & "')"
        Conn.Execute KeluarCus
    End If
        
        
        
    ADO.Recordset.MoveFirst
    Do While Not ADO.Recordset.EOF And ADO.Recordset!Kode <> vbNullString
        Dim simpan2 As String
        simpan2 = "insert into DetailMinta1(nomorMnt,KODEBRG,stok,QTYMnt,dikirim,ket) values " & _
        "('" & NomorMnt & "','" & ADO.Recordset!Kode & "','" & ADO.Recordset!stok & "','" & ADO.Recordset!qtymnt & "','" & ADO.Recordset!dikirim & "','" & ADO.Recordset!ket & "')"
        Conn.Execute simpan2
    ADO.Recordset.MoveNext
    Loop
    
    
    ADO.Recordset.MoveFirst
    Do While Not ADO.Recordset.EOF And ADO.Recordset!Kode <> vbNullString
        Dim SimpanDetailKeluar As String
        SimpanDetailKeluar = "insert into DetailKeluar(nomorKlr,KODEBRG,stok,QTYMnt,dikirim,ket) values " & _
        "('" & NomorKLr & "','" & ADO.Recordset!Kode & "','" & ADO.Recordset!stok & "','" & ADO.Recordset!qtymnt & "','" & ADO.Recordset!dikirim & "','" & ADO.Recordset!ket & "')"
        Conn.Execute SimpanDetailKeluar
    ADO.Recordset.MoveNext
    Loop
    
    
    ADO.Recordset.MoveFirst
    Do While Not ADO.Recordset.EOF
        If ADO.Recordset!Kode <> vbNullString Then
            Call Koneksi
            RSBarang.Open "Select * from Barang where Kodebrg='" & ADO.Recordset!Kode & "'", Conn
            If Not RSBarang.EOF Then
                Dim TambahBarang1 As String
                TambahBarang1 = "update barang set jumlahbrg='" & RSBarang!jumlahbrg - ADO.Recordset!dikirim & "' where kodebrg='" & ADO.Recordset!Kode & "'"
                Conn.Execute (TambahBarang1)
            End If
        End If
    ADO.Recordset.MoveNext
    Loop
        
    Form_Activate
    Call Kosongkan
    Call TabelKosong
Else
    Form_Activate
    Call Kosongkan
    Call TabelKosong
End If
End Sub

Sub Kosongkan()
Call KosongkanDpt
Call KosongkanCus
TotalMnt = ""
TotalKrm = ""
LblKet = ""
End Sub

Sub CetakPermintaan1()
    CR.ReportFileName = App.Path & "\Permintaan1.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

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
        ADO.Recordset!Nomor = i
        ADO.Recordset.Update
    Next i
    ADO.Recordset.MoveFirst
    DG.Col = 1
End If
End Function

Private Sub NotaCus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If NotaCus = "" Then
        NotaCus = "Kosong"
        DG.SetFocus
        DG.Col = 1
    Else
        DG.SetFocus
        DG.Col = 1
    End If
End If

End Sub

Private Sub NotaDpt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If NotaDpt = "" Then
        NotaDpt = "Kosong"
        DG.SetFocus
        DG.Col = 1
    Else
        DG.SetFocus
        DG.Col = 1
    End If
End If
End Sub

Sub Bukadpt()
Combo1.Enabled = True
NamaDpt.Enabled = False
PersonDpt.Enabled = False
NotaDpt.Enabled = True
End Sub

Sub Tutupdpt()
Combo1.Enabled = False
NamaDpt.Enabled = False
PersonDpt.Enabled = False
NotaDpt.Enabled = False
End Sub

Sub BukaCus()
Combo2.Enabled = True
NamaCus.Enabled = False
PersonCus.Enabled = False
NotaCus.Enabled = True
End Sub

Sub TutupCus()
Combo2.Enabled = False
NamaCus.Enabled = False
PersonCus.Enabled = False
NotaCus.Enabled = False
End Sub


Sub KondisiAwal()
'Form_Activate
Call Tutupdpt
Call TutupCus
Call KosongkanDpt
Call KosongkanCus
Call Kosongkan
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
    Call TutupCus
    Call KosongkanCus
    Call Bukadpt
    Combo1.Enabled = True
    Combo1.SetFocus
End If
End Sub


Private Sub Option2_Click()
If Option2.Value = True Then
    Call BukaCus
    Call KosongkanDpt
    Call Tutupdpt
    Combo2.SetFocus
End If
End Sub

Sub KosongkanDpt()
Combo1 = ""
NamaDpt = ""
PersonDpt = ""
NotaDpt = ""
End Sub

Sub KosongkanCus()
Combo2 = ""
NamaCus = ""
PersonCus = ""
NotaCus = ""
End Sub

