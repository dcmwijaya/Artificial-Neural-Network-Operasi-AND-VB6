VERSION 5.00
Begin VB.Form Beranda 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "APLIKASI JST-AND"
   ClientHeight    =   6570
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "Curlz MT"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Beranda.frx":0000
   LinkTopic       =   "Beranda"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Beranda.frx":0A72
   ScaleHeight     =   12306.77
   ScaleMode       =   0  'User
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Created at 2021 | FP PENGENALAN POLA-C"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   840
      TabIndex        =   3
      Top             =   5400
      Width           =   8055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Devan Cakra M. W (18081010013)"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   3000
      Width           =   8055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MK    POLA    (JST    -    AND)"
      BeginProperty Font 
         Name            =   "Poplar Std"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   2040
      Width           =   8055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "APLIKASI     FINAL    PROJECT"
      BeginProperty Font 
         Name            =   "Poplar Std"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   8055
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

Private Sub Label1_Click()

End Sub

Private Sub Run_Click(Index As Integer)
    JSTand.Show
    Beranda.Hide
    Unload Me
End Sub
