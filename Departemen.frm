VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Departemen 
   Caption         =   "Data Departemen"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6375
   LinkTopic       =   "Form2"
   ScaleHeight     =   4560
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Height          =   350
      Left            =   1680
      TabIndex        =   11
      Top             =   1920
      Width           =   3500
   End
   Begin VB.TextBox Text5 
      Height          =   350
      Left            =   1680
      TabIndex        =   10
      Top             =   1560
      Width           =   2000
   End
   Begin VB.TextBox Text4 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   1680
      TabIndex        =   9
      Top             =   1200
      Width           =   1995
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   5400
      TabIndex        =   5
      Top             =   1920
      Width           =   850
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   350
      Left            =   5400
      TabIndex        =   4
      Top             =   1560
      Width           =   850
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   350
      Left            =   5400
      TabIndex        =   3
      Top             =   1200
      Width           =   850
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   350
      Left            =   5400
      TabIndex        =   2
      Top             =   840
      Width           =   850
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   350
      Left            =   5400
      TabIndex        =   1
      Top             =   480
      Width           =   850
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   350
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   850
   End
   Begin VB.TextBox Text3 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   1680
      TabIndex        =   8
      Top             =   840
      Width           =   2000
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1680
      TabIndex        =   7
      Top             =   480
      Width           =   3500
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   2000
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   375
      Left            =   120
      Top             =   5160
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Ado"
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
      Bindings        =   "Departemen.frx":0000
      Height          =   2055
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "KodeDpt"
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
      BeginProperty Column01 
         DataField       =   "NamaDpt"
         Caption         =   "Nama Departemen"
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
         DataField       =   "PersonDpt"
         Caption         =   "Person"
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
         DataField       =   "Telepon"
         Caption         =   "Telepon"
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
         DataField       =   "HP"
         Caption         =   "HP"
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
         DataField       =   "Email"
         Caption         =   "Email"
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
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telepon"
      Height          =   345
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " HP"
      Height          =   345
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " E-Mail"
      Height          =   345
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Contact Person"
      Height          =   345
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Departemen"
      Height          =   345
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Departemen"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "Departemen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Call Koneksi
ADO.ConnectionString = PathData
ADO.RecordSource = "Departemen"
ADO.Refresh
Set DG.DataSource = ADO
DG.Refresh
End Sub

Sub Form_Load()
    Call Koneksi
    Text1.MaxLength = 3
    Text2.MaxLength = 30
    Text3.MaxLength = 30
    Text4.MaxLength = 15
    Text5.MaxLength = 15
    Text6.MaxLength = 30
    KondisiAwal
End Sub


Private Sub dg_Keypress(Keyascii As Integer)
If Keyascii = 13 Then
    If CmdSimpan.Enabled = False Then
        MsgBox "pilih dulu command Edit atau Hapus"
        Exit Sub
    End If
    If CmdEdit.Enabled = True Then
        Text1.Enabled = False
        Text1 = DG.Columns(0)
        Text2 = DG.Columns(1)
        Text3 = DG.Columns(2)
        Text4 = DG.Columns(3)
        Text5 = DG.Columns(4)
        Text6 = DG.Columns(5)
        Text2.SetFocus
    End If
    
    If CmdHapus.Enabled = True Then
        Text1 = DG.Columns(0)
        Text2 = DG.Columns(1)
        Text3 = DG.Columns(2)
        Text4 = DG.Columns(3)
        Text5 = DG.Columns(4)
        Text6 = DG.Columns(5)
        Call CariData
        If Not RSDepartemen.EOF Then
            Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
            If Pesan = vbYes Then
                hapus = "delete * from Departemen where kodeDpt='" & Text1 & "'"
                Conn.Execute hapus
                Call KondisiAwal
            Else
                Call KondisiAwal
            End If
        End If
    End If
End If
End Sub

Private Sub KODEOTO()
Call Koneksi
RSDepartemen.Open ("select * from Departemen Where KodeDpt In(Select Max(KodeDpt)From Departemen)Order By KodeDpt Desc"), Conn
RSDepartemen.Requery
    Dim Urutan As String * 3
    Dim Hitung As Long
    With RSDepartemen
        If .EOF Then
            Urutan = "D" + "01"
            Text1 = Urutan
        Else
            Hitung = Right(!Kodedpt, 2) + 1
            Urutan = "D" + Right("00" & Hitung, 2)
        End If
        Text1 = Urutan
    End With

End Sub

Function CariData()
    Call Koneksi
    RSDepartemen.Open "Select * From Departemen where KodeDpt='" & Text1 & "'", Conn
End Function

Private Sub CmdBatal_Click()
KosongkanText
TidakSiapIsi
KondisiAwal
End Sub

Private Sub CmdSimpan_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
    MsgBox "Data Belum Lengkap...!"
    Exit Sub
Else
    If CmdInput.Enabled = True Then
            Dim SQLTambah1 As String
            SQLTambah1 = "Insert Into Departemen (KodeDpt,NamaDpt,PersonDpt,telepon,HP,email) values " & _
            "('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "','" & Text6 & "')"
            Conn.Execute SQLTambah1
    Else
            Dim SQLEdit As String
            SQLEdit = "Update Departemen Set NamaDpt= '" & Text2 & "', PersonDpt = '" & Text3 & "', telepon= '" & Text4 & "',HP= '" & Text5 & "', email= '" & Text6 & "' where KodeDpt='" & Text1 & "'"
            Conn.Execute SQLEdit
        End If
    Form_Activate
    KondisiAwal
End If
End Sub

Private Sub KosongkanText()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
End Sub

Private Sub SiapIsi()
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    
End Sub

Private Sub TidakSiapIsi()
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Text6.Enabled = False
End Sub

Private Sub KondisiAwal()
KosongkanText
TidakSiapIsi
CmdInput.Enabled = True
CmdEdit.Enabled = True
CmdHapus.Enabled = True
CmdSimpan.Enabled = False
CmdBatal.Enabled = False
CmdTutup.Enabled = True
Form_Activate
End Sub

Private Sub TampilkanData()
On Error Resume Next
Text2 = RSDepartemen!NamaDpt
Text3 = RSDepartemen!PersonDpt
Text4 = RSDepartemen!telepon
Text5 = RSDepartemen!HP
Text6 = RSDepartemen!email
End Sub

Private Sub CmdInput_Click()
    If CmdInput.Caption = "&Input" Then
        CmdEdit.Enabled = False
        CmdHapus.Enabled = False
        CmdSimpan.Enabled = True
        CmdBatal.Enabled = True
        CmdTutup.Enabled = False
        SiapIsi
        KosongkanText
        Call KODEOTO
        Text1.Enabled = False
        Text2.SetFocus
    End If
End Sub

Private Sub CmdEdit_Click()
If CmdEdit.Caption = "&Edit" Then
    CmdInput.Enabled = False
    CmdHapus.Enabled = False
    CmdTutup.Enabled = False
    CmdSimpan.Enabled = True
    CmdBatal.Enabled = True
    SiapIsi
    Text1.SetFocus
End If
End Sub

Private Sub CmdHapus_Click()
If CmdHapus.Caption = "&Hapus" Then
    CmdTutup.Enabled = False
    CmdInput.Enabled = False
    CmdEdit.Enabled = False
    CmdSimpan.Enabled = True
    CmdBatal.Enabled = True
    SiapIsi
    Text1.SetFocus
End If

End Sub

Private Sub CmdTutup_Click()
    Select Case CmdTutup.Caption
        Case "&Tutup"
            Unload Me
        Case "&Batal"
            TidakSiapIsi
            KondisiAwal
    End Select
End Sub

Private Sub Text1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If CmdInput.Enabled = True Then
        Call CariData
            If Not RSDepartemen.EOF Then
                TampilkanData
                MsgBox "Kode Departemen Sudah Ada"
                KosongkanText
                Text1.SetFocus
            Else
                Text2.SetFocus
            End If
    End If
    
    If CmdEdit.Enabled = True Then
        Call CariData
            If Not RSDepartemen.EOF Then
                TampilkanData
                Text1.Enabled = False
                Text2.SetFocus
            Else
                MsgBox "Kode Departemen Tidak Ada"
                Text1 = ""
                Text1.SetFocus
            End If
    End If
    
    If CmdHapus.Enabled = True Then
        Call CariData
            If Not RSDepartemen.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From Departemen where KodeDpt= '" & Text1 & "'"
                    Conn.Execute SQLHapus
                    Form_Activate
                    KondisiAwal
                Else
                    KondisiAwal
                    CmdHapus.SetFocus
                End If
            Else
                MsgBox "Data Tidak ditemukan"
                Text1.SetFocus
            End If
    End If
End If
End Sub

Private Sub text2_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text3.SetFocus
End Sub

Private Sub text3_keypress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text4.SetFocus
End Sub

Private Sub text4_keypress(Keyascii As Integer)
    If Keyascii = 13 Then Text5.SetFocus
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub text5_keypress(Keyascii As Integer)
    If Keyascii = 13 Then Text6.SetFocus
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub text6_keypress(Keyascii As Integer)

    If Keyascii = 13 Then
        If CmdInput.Enabled = True Then
            CmdSimpan.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdSimpan.SetFocus
        End If
    End If
End Sub

