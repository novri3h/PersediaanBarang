VERSION 5.00
Begin VB.Form Login 
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3930
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
   ScaleHeight     =   2160
   ScaleWidth      =   3930
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtKodePmk 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   1440
      TabIndex        =   5
      Top             =   2520
      Width           =   2025
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1155
      ScaleWidth      =   3345
      TabIndex        =   2
      Top             =   720
      Width           =   3400
      Begin VB.TextBox TxtPasswordPmk 
         Height          =   350
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "X"
         TabIndex        =   1
         Top             =   600
         Width           =   2000
      End
      Begin VB.TextBox TxtNamaPmk 
         Height          =   350
         IMEMode         =   3  'DISABLE
         Left            =   1200
         TabIndex        =   0
         Top             =   120
         Width           =   2000
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Password"
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nama User"
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1005
      End
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Kasir"
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
      Left            =   360
      TabIndex        =   7
      Top             =   2520
      Width           =   1005
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim A As Byte
Dim B As Byte

Private Sub Form_Load()
    TxtNamaPmk.MaxLength = 35
    TxtPasswordPmk.MaxLength = 15
    TxtPasswordPmk.PasswordChar = "*"
    TxtPasswordPmk.Enabled = False
    TxtKodePmk.Enabled = False
End Sub

Private Sub TxtNamaPmk_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 27 Then Unload Me
If Keyascii = 13 Then
   
    Call Koneksi
    RSPemakai.Open "Select NamaPmk from Pemakai where NamaPmk ='" & TxtNamaPmk & "'", Conn
    If RSPemakai.EOF Then
        A = A + 1
        If 1 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaPmk & "' tidak dikenal"
            TxtNamaPmk = ""
            TxtNamaPmk.SetFocus
        ElseIf 2 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaPmk & "' tidak dikenal"
            TxtNamaPmk = ""
            TxtNamaPmk.SetFocus
        ElseIf 3 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaPmk & "' tidak dikenal" & Chr(13) & _
                    "Kesempatan habis, Ulangi dari awal"
            Conn.Close
            End
            'Conn.Close
            'Unload Me
        End If
    Else
        TxtNamaPmk.Enabled = False
        TxtPasswordPmk.Enabled = True
        TxtPasswordPmk.SetFocus
        Conn.Close
    End If
End If
End Sub

Private Sub txtpasswordPmk_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 27 Then Unload Me
If Keyascii = 13 Then

    Call Koneksi
    RSPemakai.Open "Select * from Pemakai where NamaPmk ='" & TxtNamaPmk & "' and PassPmk='" & TxtPasswordPmk & "'", Conn
    If RSPemakai.EOF Then
        B = B + 1
        If 1 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            TxtPasswordPmk = ""
            TxtPasswordPmk.SetFocus
        ElseIf 2 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            TxtPasswordPmk = ""
            TxtPasswordPmk.SetFocus
        ElseIf 3 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            'End
            Conn.Close
            End
            'Unload Me
        End If
    Else
        Unload Me
        Menu.Show
        Menu.STBar.Panels(1).Text = RSPemakai!KodePmk
        Menu.STBar.Panels(2).Text = RSPemakai!NamaPmk
        Menu.STBar.Panels(3).Text = RSPemakai!StatusPmk
        If Menu.STBar.Panels(3).Text <> "ADMINISTRATOR" Then
            Menu.mnfile.Enabled = False
        End If
        Conn.Close
    End If
End If
End Sub



