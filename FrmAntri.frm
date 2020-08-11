VERSION 5.00
Begin VB.Form FrmAntri 
   Caption         =   "Aplikasi Antrian"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7260
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAntri.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmAntri.frx":57E2
   ScaleHeight     =   3060
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton op3 
      Caption         =   "CS04"
      Height          =   255
      Left            =   5640
      TabIndex        =   20
      Top             =   240
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "CS03"
      Height          =   255
      Left            =   3720
      TabIndex        =   19
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   18
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   17
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   16
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   15
      Top             =   2520
      Width           =   375
   End
   Begin VB.OptionButton op2 
      Caption         =   "CS01"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin VB.OptionButton op1 
      Caption         =   "CS02"
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Loket 4"
      Height          =   495
      Left            =   5400
      TabIndex        =   12
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox TA4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5400
      TabIndex        =   11
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox TH4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Loket 3"
      Height          =   495
      Left            =   3600
      TabIndex        =   9
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox TA3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3600
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox TH3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox thuruf 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3600
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox TH2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TA2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1800
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Loket 2"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox TH1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TA1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Loket 1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   1455
   End
End
Attribute VB_Name = "FrmAntri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim no As Integer
Dim Sounds(16), alamat As String

Sub Bunyi()
    Dim kar() As String
    kar = Split(thuruf.Text, " ")
    Call sndPlaySound(alamat & "nomor-urut.wav", SND_NOSTOP)
    For i = LBound(kar) To UBound(kar)
        Call sndPlaySound(alamat & kar(i) & ".wav", SND_NOSTOP)
    Next
    Call sndPlaySound(alamat & "loket.wav", SND_NOSTOP)
End Sub

Sub Ulang(nomor As String, nom As String)
    Dim kar() As String
    kar = Split(nom, " ")
    Call sndPlaySound(alamat & "nomor-urut.wav", SND_NOSTOP)
    For i = LBound(kar) To UBound(kar)
        Call sndPlaySound(alamat & kar(i) & ".wav", SND_NOSTOP)
    Next
    Call sndPlaySound(alamat & "loket.wav", SND_NOSTOP)
    Call sndPlaySound(alamat & nomor & ".wav", SND_NOSTOP)
End Sub

Sub Loket1()
  Call Bunyi
  Call sndPlaySound(alamat & "satu.wav", SND_NOSTOP)
End Sub

Sub Loket2()
  Call Bunyi
  Call sndPlaySound(alamat & "dua.wav", SND_NOSTOP)
End Sub

Sub Loket3()
  Call Bunyi
  Call sndPlaySound(alamat & "tiga.wav", SND_NOSTOP)
End Sub

Sub Loket4()
  Call Bunyi
  Call sndPlaySound(alamat & "empat.wav", SND_NOSTOP)
End Sub

Private Sub Command1_Click()
    no = no + 1
    TA1.Text = no
    thuruf.Text = Trim(Bilang(TA1.Text))
    TH1.Text = Trim(Bilang(TA1.Text))
    Call Loket1
End Sub

Private Sub Command2_Click()
    no = no + 1
    TA2.Text = no
    thuruf.Text = Trim(Bilang(TA2.Text))
    TH2.Text = Trim(Bilang(TA2.Text))
    Call Loket2
End Sub

Private Sub Command3_Click()
    no = no + 1
    TA3.Text = no
    thuruf.Text = Trim(Bilang(TA3.Text))
    TH3.Text = Trim(Bilang(TA3.Text))
    Call Loket3
End Sub

Private Sub Command4_Click()
    no = no + 1
    TA4.Text = no
    thuruf.Text = Trim(Bilang(TA4.Text))
    TH4.Text = Trim(Bilang(TA4.Text))
    Call Loket4
End Sub

Private Sub Command5_Click()
 Ulang "satu", TH1.Text
End Sub

Private Sub Command6_Click()
 Ulang "dua", TH2.Text
End Sub

Private Sub Command7_Click()
  Ulang "tiga", TH3.Text
End Sub

Private Sub Command8_Click()
 Ulang "empat", TH4.Text
End Sub

Private Sub Form_Load()
   no = 0
   alamat = App.Path & "\sounds\ladies\"
   Sounds(1) = alamat & "satu.wav"
   Sounds(2) = alamat & "dua.wav"
   Sounds(3) = alamat & "tiga.wav"
   Sounds(4) = alamat & "empat.wav"
   Sounds(5) = alamat & "lima.wav"
   Sounds(6) = alamat & "enam.wav"
   Sounds(7) = alamat & "tujuh.wav"
   Sounds(8) = alamat & "delapan.wav"
   Sounds(9) = alamat & "sembilan.wav"
   Sounds(10) = alamat & "sepuluh.wav"
   Sounds(11) = alamat & "sebelas.wav"
   Sounds(12) = alamat & "puluh.wav"
   Sounds(13) = alamat & "ratus.wav"
   Sounds(14) = alamat & "belas.wav"
   Sounds(15) = alamat & "nomor-urut.wav"
   Sounds(16) = alamat & "loket.wav"
End Sub

Private Sub op1_Click()
  alamat = App.Path & "\sounds\cs01\"
End Sub

Private Sub op2_Click()
  alamat = App.Path & "\sounds\cs02\"
End Sub

Private Sub Option1_Click()
  alamat = App.Path & "\sounds\cs03\"
End Sub
Private Sub op3_Click()
  alamat = App.Path & "\sounds\cs04\"
End Sub

