VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form LapTransaksi 
   Caption         =   "Laporan Transaksi"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5835
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
   ScaleHeight     =   4230
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Bantuan"
      Height          =   375
      Left            =   2280
      TabIndex        =   26
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   3960
      TabIndex        =   25
      Top             =   3600
      Width           =   1575
   End
   Begin Crystal.CrystalReport CR 
      Left            =   240
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   5953
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Permintaan"
      TabPicture(0)   =   "LapTransaksi.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Combo1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Combo2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Combo3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Penerimaan"
      TabPicture(1)   =   "LapTransaksi.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label8"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label9"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label14"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Combo4"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Combo6"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Combo5"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Pengeluaran"
      TabPicture(2)   =   "LapTransaksi.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label10"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label11"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label12"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label13"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label15"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Combo7"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Combo8"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Combo9"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      Begin VB.ComboBox Combo5 
         Height          =   345
         Left            =   -72240
         TabIndex        =   24
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox Combo9 
         Height          =   345
         Left            =   -71160
         TabIndex        =   7
         Top             =   2160
         Width           =   1455
      End
      Begin VB.ComboBox Combo8 
         Height          =   345
         Left            =   -71160
         TabIndex        =   6
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox Combo7 
         Height          =   345
         Left            =   -71160
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox Combo6 
         Height          =   345
         Left            =   -72240
         TabIndex        =   4
         Top             =   2160
         Width           =   1455
      End
      Begin VB.ComboBox Combo4 
         Height          =   345
         Left            =   -72240
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox Combo3 
         Height          =   345
         Left            =   1320
         TabIndex        =   2
         Top             =   2160
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   345
         Left            =   1320
         TabIndex        =   1
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   1320
         TabIndex        =   0
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Laporan Bulanan"
         Height          =   225
         Left            =   -72240
         TabIndex        =   23
         Top             =   1440
         Width           =   1350
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Laporan Bulanan"
         Height          =   225
         Left            =   -73320
         TabIndex        =   22
         Top             =   1440
         Width           =   1350
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal"
         Height          =   345
         Left            =   -72240
         TabIndex        =   21
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bulan"
         Height          =   345
         Left            =   -72240
         TabIndex        =   20
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tahun"
         Height          =   345
         Left            =   -72240
         TabIndex        =   19
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Laporan Harian"
         Height          =   225
         Left            =   -72240
         TabIndex        =   18
         Top             =   480
         Width           =   1230
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal"
         Height          =   345
         Left            =   -73320
         TabIndex        =   17
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bulan"
         Height          =   345
         Left            =   -73320
         TabIndex        =   16
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tahun"
         Height          =   345
         Left            =   -73320
         TabIndex        =   15
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Laporan Harian"
         Height          =   225
         Left            =   -73320
         TabIndex        =   14
         Top             =   480
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Laporan Bulanan"
         Height          =   225
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Laporan Harian"
         Height          =   225
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   1230
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tahun"
         Height          =   345
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bulan"
         Height          =   345
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal"
         Height          =   345
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1005
      End
   End
End
Attribute VB_Name = "LapTransaksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'On Error Resume Next
Call Koneksi
RSPermintaanUser.Open "Select Distinct TanggalMnt From Permintaanuser order By 1", Conn
RSPermintaanUser.Requery
Do Until RSPermintaanUser.EOF
    Combo1.AddItem Format(RSPermintaanUser!TanggalMnt, "DD-MMM-YYYY")
    RSPermintaanUser.MoveNext
Loop

Dim RSTGLMinta As New ADODB.Recordset
RSTGLMinta.Open "select distinct month(TanggalMnt) as Bulan from Permintaanuser", Conn
Do While Not RSTGLMinta.EOF
    Combo2.AddItem RSTGLMinta!Bulan & Space(5) & MonthName(RSTGLMinta!Bulan)
    RSTGLMinta.MoveNext
Loop

Dim RSThnMinta As New ADODB.Recordset
RSThnMinta.Open "select distinct year(TanggalMnt)  as Tahun from Permintaanuser", Conn
Do While Not RSThnMinta.EOF
    Combo3.AddItem RSThnMinta!Tahun
    RSThnMinta.MoveNext
Loop

RSPenerimaan.Open "Select Distinct TanggalTrm From Penerimaan order By 1", Conn
RSPenerimaan.Requery
Do Until RSPenerimaan.EOF
    Combo4.AddItem Format(RSPenerimaan!TanggalTrm, "DD-MMM-YYYY")
    RSPenerimaan.MoveNext
Loop

Dim RSTGLTerima As New ADODB.Recordset
RSTGLTerima.Open "select distinct month(TanggalTrm) as Bulan from Penerimaan", Conn
Do While Not RSTGLTerima.EOF
    Combo5.AddItem RSTGLTerima!Bulan & Space(5) & MonthName(RSTGLTerima!Bulan)
    RSTGLTerima.MoveNext
Loop

Dim RSTHNTerima As New ADODB.Recordset
RSTHNTerima.Open "select distinct year(TanggalTrm)  as Tahun from Penerimaan", Conn
Do While Not RSTHNTerima.EOF
    Combo6.AddItem RSTHNTerima!Tahun
    RSTHNTerima.MoveNext
Loop


RSPengeluaran.Open "Select Distinct TanggalKlr From Pengeluaran order By 1", Conn
RSPengeluaran.Requery
Do Until RSPengeluaran.EOF
    Combo7.AddItem Format(RSPengeluaran!TanggalKlr, "DD-MMM-YYYY")
    RSPengeluaran.MoveNext
Loop

Dim RSTGLKeluar As New ADODB.Recordset
RSTGLKeluar.Open "select distinct month(TanggalTrm) as Bulan from Penerimaan", Conn
Do While Not RSTGLKeluar.EOF
    Combo8.AddItem RSTGLKeluar!Bulan & Space(5) & MonthName(RSTGLKeluar!Bulan)
    RSTGLKeluar.MoveNext
Loop

Dim RSTHNKeluar As New ADODB.Recordset
RSTHNKeluar.Open "select distinct year(TanggalTrm)  as Tahun from Penerimaan", Conn
Do While Not RSTHNKeluar.EOF
    Combo9.AddItem RSTHNKeluar!Tahun
    RSTHNKeluar.MoveNext
Loop

Conn.Close
End Sub

Private Sub COMBO1_Click()
    CR.SelectionFormula = "Totext({Permintaanuser.TanggalMnt})='" & CDate(Combo1) & "'"
    CR.ReportFileName = App.Path & "\lap minta user harian.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Combo3_Click()
    Call Koneksi
    RSPermintaanUser.Open "select * from Permintaanuser where month(TanggalMnt)='" & Val(Left(Combo2, 2)) & "' and year(TanggalMnt)='" & (Combo3) & "'", Conn
    If RSPermintaanUser.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo4.SetFocus
    End If
    CR.SelectionFormula = "Month({Permintaanuser.TanggalMnt})=" & Val(Left(Combo2, 2)) & " and Year({Permintaanuser.TanggalMnt})=" & Val(Combo3.Text)
    CR.ReportFileName = App.Path & "\LAP MINTA USER BULANAN.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Combo4_Click()
    CR.SelectionFormula = "Totext({Penerimaan.Tanggaltrm})='" & CDate(Combo4) & "'"
    CR.ReportFileName = App.Path & "\LAP TERIMA HARIAN.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Combo6_Click()
Call Koneksi
RSPenerimaan.Open "select * from PENERIMAAN where month(TanggalTRM)='" & Val(Left(Combo5, 2)) & "' and year(TanggalTRM)='" & (Combo6) & "'", Conn
If RSPenerimaan.EOF Then
    MsgBox "Data tidak ditemukan"
    Exit Sub
    Combo4.SetFocus
End If
CR.SelectionFormula = "Month({PENERIMAAN.TanggalTRM})=" & Val(Left(Combo5, 2)) & " and Year({PENERIMAAN.TanggalTRM})=" & Val(Combo6.Text)
CR.ReportFileName = App.Path & "\LAP TERIMA BULANAN.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1

End Sub

Private Sub Combo7_Click()
    CR.SelectionFormula = "Totext({pengeluaran.TanggalKlr})='" & CDate(Combo7) & "'"
    CR.ReportFileName = App.Path & "\LAP keluar HARIAN.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Combo9_Click()
Call Koneksi
RSPengeluaran.Open "select * from PENGELUARAN where month(TanggalKLR)='" & Val(Left(Combo8, 2)) & "' and year(TanggalKLR)='" & (Combo9) & "'", Conn
If RSPengeluaran.EOF Then
    MsgBox "Data tidak ditemukan"
    Exit Sub
    Combo4.SetFocus
End If
CR.SelectionFormula = "Month({PENGELUARAN.TanggalKLR})=" & Val(Left(Combo8, 2)) & " and Year({PENGELUARAN.TanggalKLR})=" & Val(Combo9.Text)
CR.ReportFileName = App.Path & "\LAP KELUAR BULANAN.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1

End Sub


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
MsgBox "Untuk laporan harian" & Chr(13) & _
        "Pilih salah satu tanggal dalam combo" & Chr(13) & _
        "" & Chr(13) & _
        "untuk laporan bulanan" & Chr(13) & _
        "pilih bulan kemudian pilih tahun"
End Sub

