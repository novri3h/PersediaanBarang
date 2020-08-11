VERSION 5.00
Begin VB.Form KodeBarang 
   Caption         =   "Kode Barang (ESC = Tutup)"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   4140
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "KodeBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call Koneksi
RSBarang.Open "barang", Conn
List1.Clear
Do While Not RSBarang.EOF
    List1.AddItem RSBarang!kodebrg & vbTab & RSBarang!namabrg
    RSBarang.MoveNext
Loop
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
