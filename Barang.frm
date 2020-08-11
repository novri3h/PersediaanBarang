VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Barang 
   Caption         =   "Data Barang"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5310
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
   ScaleHeight     =   5565
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   1680
      TabIndex        =   17
      Top             =   1920
      Width           =   2000
   End
   Begin VB.TextBox Text4 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   1680
      TabIndex        =   15
      Top             =   1560
      Width           =   2000
   End
   Begin MSDataGridLib.DataGrid DG 
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   18
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
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1680
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   840
      Width           =   2000
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   375
      Left            =   120
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   2000
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1680
      TabIndex        =   7
      Top             =   480
      Width           =   3540
   End
   Begin VB.TextBox Text3 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   1680
      TabIndex        =   8
      Top             =   1200
      Width           =   2000
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   400
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   850
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   400
      Left            =   960
      TabIndex        =   1
      Top             =   2400
      Width           =   850
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   400
      Left            =   1800
      TabIndex        =   2
      Top             =   2400
      Width           =   850
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   400
      Left            =   2640
      TabIndex        =   3
      Top             =   2400
      Width           =   850
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   400
      Left            =   3480
      TabIndex        =   4
      Top             =   2400
      Width           =   850
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   400
      Left            =   4320
      TabIndex        =   5
      Top             =   2400
      Width           =   850
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jumlah Stok"
      Height          =   345
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Stok Maksimum"
      Height          =   345
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Satuan"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Barang"
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Barang"
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Stok Minimum"
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1500
   End
End
Attribute VB_Name = "Barang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Call Koneksi
ADO.ConnectionString = PathData
ADO.RecordSource = "Barang"
ADO.Refresh
Set DG.DataSource = ADO
DG.Refresh
RSBarang.Open "SELECT DISTINCT SATUAN FROM barang", Conn
Combo1.Clear
Do While Not RSBarang.EOF
    Combo1.AddItem RSBarang!satuan
    RSBarang.MoveNext
Loop
Conn.Close
End Sub

Sub Form_Load()
    Call Koneksi
    Text1.MaxLength = 3
    Text2.MaxLength = 30
    Text3.MaxLength = 5
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
        Combo1 = DG.Columns(2)
        Text3 = DG.Columns(3)
        Text4 = DG.Columns(4)
        Text5 = DG.Columns(5)
        Text2.SetFocus
    End If
    
    If CmdHapus.Enabled = True Then
        Text1 = DG.Columns(0)
        Text2 = DG.Columns(1)
        Combo1 = DG.Columns(2)
        Text3 = DG.Columns(3)
        Text4 = DG.Columns(4)
        Text5 = DG.Columns(5)
        Call CariData
        If Not RSBarang.EOF Then
            Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
            If Pesan = vbYes Then
                hapus = "delete * from Barang where kodebrg='" & Text1 & "'"
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
RSBarang.Open ("select * from Barang Where KodeBrg In(Select Max(KodeBrg)From Barang)Order By KodeBrg Desc"), Conn
RSBarang.Requery
    Dim Urutan As String * 3
    Dim Hitung As Long
    With RSBarang
        If .EOF Then
            Urutan = "001"
            Text1 = Urutan
        Else
            Hitung = !kodebrg + 1
            Urutan = Right("000" & Hitung, 3)
        End If
        Text1 = Urutan
    End With

End Sub

Function CariData()
    Call Koneksi
    RSBarang.Open "Select * From Barang where KodeBrg='" & Text1 & "'", Conn
End Function

Private Sub CmdBatal_Click()
KosongkanText
TidakSiapIsi
KondisiAwal
End Sub

Private Sub CmdSimpan_Click()
If Text1 = "" Or Text2 = "" Or Combo1 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
    MsgBox "Data Belum Lengkap...!"
    Exit Sub
Else
    Call Koneksi
    If CmdInput.Enabled = True Then
            Dim SQLTambah1 As String
            SQLTambah1 = "Insert Into Barang (KodeBrg,NamaBrg,Satuan,stokMinimum,StokMaksimum,JumlahBrg) values " & _
            "('" & Text1 & "','" & Text2 & "','" & Combo1 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "')"
            Conn.Execute SQLTambah1
    Else
            Dim SQLEdit As String
            SQLEdit = "Update Barang Set NamaBrg= '" & Text2 & "', satuan='" & Combo1 & "',stokMinimum= '" & Text3 & "',stokMaksimum= '" & Text4 & "',JumlahBrg = '" & Text5 & "' where KodeBrg='" & Text1 & "'"
            Conn.Execute SQLEdit
        End If
    Form_Activate
    KondisiAwal
End If
End Sub

Private Sub KosongkanText()
    Text1 = ""
    Text2 = ""
    Combo1 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
End Sub

Private Sub SiapIsi()
    Text1.Enabled = True
    Text2.Enabled = True
    Combo1.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Text1.Enabled = False
    Text2.Enabled = False
    Combo1.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
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
Text2 = RSBarang!namabrg
Combo1 = RSBarang!satuan
Text3 = RSBarang!StokMinimum
Text4 = RSBarang!stokmaksimum
Text5 = RSBarang!jumlahbrg
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
            If Not RSBarang.EOF Then
                TampilkanData
                MsgBox "Kode Barang Sudah Ada"
                KosongkanText
                Text1.SetFocus
            Else
                Text2.SetFocus
            End If
    End If
    
    If CmdEdit.Enabled = True Then
        Call CariData
            If Not RSBarang.EOF Then
                TampilkanData
                Text1.Enabled = False
                Text2.SetFocus
            Else
                MsgBox "Kode Barang Tidak Ada"
                Text1 = ""
                Text1.SetFocus
            End If
    End If
    
    If CmdHapus.Enabled = True Then
        Call CariData
            If Not RSBarang.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From Barang where KodeBrg= '" & Text1 & "'"
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
    If Keyascii = 13 Then Combo1.SetFocus
End Sub

Private Sub COMBO1_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text3.SetFocus
End Sub

Private Sub text3_keypress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text4.SetFocus
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub text4_keypress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text5.SetFocus
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub


Private Sub text5_keypress(Keyascii As Integer)
    If Keyascii = 13 Then
        If CmdInput.Enabled = True Then
            CmdSimpan.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdSimpan.SetFocus
        End If
    End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

