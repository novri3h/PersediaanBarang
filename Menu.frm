VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Menu 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Menu Utama Aplikasi Persediaan Barang"
   ClientHeight    =   3090
   ClientLeft      =   195
   ClientTop       =   765
   ClientWidth     =   5250
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
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CR 
      Left            =   1920
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.StatusBar STBar 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2595
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":2A8D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":2AAAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":2B2C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":2B5DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":2B8F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":2BC12
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":2BDEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":2C606
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":2CE20
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":2D63A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnfile 
      Caption         =   "File"
      Begin VB.Menu mnbarang 
         Caption         =   "Barang"
      End
      Begin VB.Menu mnpemakai 
         Caption         =   "Pemakai"
      End
      Begin VB.Menu mnsupplier 
         Caption         =   "Supplier"
      End
      Begin VB.Menu mncustomer 
         Caption         =   "Customer"
      End
   End
   Begin VB.Menu mntransaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu mnpermintaan 
         Caption         =   "Cek Stok Barang Di Gudang"
      End
      Begin VB.Menu mnterima 
         Caption         =   "Penerimaan Barang Dari Supplier"
      End
      Begin VB.Menu mnmintadariuser 
         Caption         =   "Permintaan Barang Dari User"
      End
      Begin VB.Menu mnpengeluaran 
         Caption         =   "Pengeluaran Barang Untuk User"
      End
   End
   Begin VB.Menu mnlaporan 
      Caption         =   "Laporan"
      Begin VB.Menu mnlapmaster 
         Caption         =   "Laporan Data Master"
         Begin VB.Menu mnlapbarang 
            Caption         =   "Master Barang"
         End
         Begin VB.Menu mnlapsupplier 
            Caption         =   "Master Supplier"
         End
         Begin VB.Menu mnlapcustomer 
            Caption         =   "Master Customer"
         End
         Begin VB.Menu mnlappemakai 
            Caption         =   "Master Pemakai"
         End
      End
      Begin VB.Menu mnlaptransaksi 
         Caption         =   "Laporan Transaksi"
      End
      Begin VB.Menu mnlapperfaktur 
         Caption         =   "Rincian Transaksi"
      End
      Begin VB.Menu mnlapstok 
         Caption         =   "Rincian Stok Barang Per tanggal"
      End
   End
   Begin VB.Menu mnkeluar 
      Caption         =   "Keluar"
      Begin VB.Menu mnlogout 
         Caption         =   "Log Out"
      End
      Begin VB.Menu mntutup 
         Caption         =   "Tutup Aplikasi"
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then End
End Sub

Private Sub mnbarang_Click()
Barang.Show vbModal

End Sub

Private Sub mncustomer_Click()
Customer.Show vbModal
End Sub

Private Sub mndepartemen_Click()
Departemen.Show vbModal

End Sub

Private Sub mnlapbarang_Click()
    CR.ReportFileName = App.Path & "\lap barang.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub mnlapcustomer_Click()
    CR.ReportFileName = App.Path & "\lap customer.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub mnlappemakai_Click()
    CR.ReportFileName = App.Path & "\lap pemakai.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub mnlapperfaktur_Click()
RincianTransaksi.Show vbModal
End Sub

Private Sub mnlapstok_Click()
RincianStok.Show vbModal
End Sub

Private Sub mnlapsupplier_Click()
    CR.ReportFileName = App.Path & "\lap supplier.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub mnlaptransaksi_Click()
LapTransaksi.Show vbModal
End Sub

Private Sub mnlogout_Click()
Menu.STBar.Panels(1).Text = ""
Menu.STBar.Panels(2).Text = ""
Menu.STBar.Panels(3).Text = ""
Login.Show
Login.TxtKodePmk = ""
Login.TxtNamaPmk = ""
Login.TxtPasswordPmk = ""
Login.TxtNamaPmk.SetFocus
End Sub

Private Sub mnmintadariuser_Click()
PermintaanUser.Show
End Sub

Private Sub mnpemakai_Click()
Pemakai.Show vbModal
End Sub

Private Sub mnpengeluaran_Click()
PengeluaranBarang.Show vbModal
End Sub

Private Sub mnpermintaan_Click()
'PermintaanDariGudang.Show vbModal
CekStokMinimum.Show vbModal
End Sub

Private Sub mnsupplier_Click()
Supplier.Show vbModal

End Sub

Private Sub mnterima_Click()
Penerimaan.Show vbModal
End Sub

Private Sub mntutup_Click()
Pesan = MsgBox("yakin akan keluar", vbYesNo)
If Pesan = vbYes Then End
End Sub

Private Sub mnujisql_Click()
UjiSQL.Show vbModal
End Sub
