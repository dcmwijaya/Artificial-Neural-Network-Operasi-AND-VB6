VERSION 5.00
Begin VB.Form Beranda 
   BackColor       =   &H00FFFF80&
   Caption         =   "APLIKASI JST-AND"
   ClientHeight    =   6420
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   9720
   Icon            =   "Beranda.frx":0000
   LinkTopic       =   "Beranda"
   Picture         =   "Beranda.frx":0A72
   ScaleHeight     =   6420
   ScaleMode       =   0  'User
   ScaleWidth      =   9720
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Created at 2021 | FP PENGENALAN POLA-C"
      BeginProperty Font 
         Name            =   "Anklepants"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   3
      Top             =   5880
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Devan Cakra M.W - 18081010013"
      BeginProperty Font 
         Name            =   "Anklepants"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   2
      Top             =   3480
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MK POLA (JST - AND)"
      BeginProperty Font 
         Name            =   "Anagram NF"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   1
      Top             =   2640
      Width           =   5655
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "APLIKASI FINAL PROJECT"
      BeginProperty Font 
         Name            =   "Anagram NF"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   2040
      Width           =   6255
   End
   Begin VB.Menu fl 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu Run 
         Caption         =   "Run"
         Index           =   0
      End
      Begin VB.Menu KLR 
         Caption         =   "Keluar"
         Index           =   0
      End
   End
End
Attribute VB_Name = "Beranda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub KLR_Click(Index As Integer)
    Q = MsgBox("Anda yakin akan keluar ?", vbQuestion + vbOKCancel, "System")
    If Q = vbOK Then
        Unload Me
        End
    End If
End Sub

Private Sub Run_Click(Index As Integer)
    JSTand.Show
    Beranda.Hide
    Unload Me
End Sub
