VERSION 5.00
Begin VB.Form JSTand 
   BackColor       =   &H00FFFFC0&
   Caption         =   "APLIKASI JST-AND"
   ClientHeight    =   6420
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   9420
   Icon            =   "JSTand.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "JSTand.frx":0A72
   ScaleHeight     =   6420
   ScaleMode       =   0  'User
   ScaleWidth      =   12306.77
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   5640
      Width           =   1935
   End
   Begin VB.TextBox Text48 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6525
      TabIndex        =   71
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text47 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5640
      TabIndex        =   70
      Top             =   2355
      Width           =   375
   End
   Begin VB.TextBox Text46 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   5400
      TabIndex        =   69
      Top             =   1350
      Width           =   379
   End
   Begin VB.TextBox Text45 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   5520
      TabIndex        =   68
      Top             =   690
      Width           =   379
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFF80&
      Caption         =   "UPDATED WEIGHT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   6480
      TabIndex        =   34
      Top             =   3240
      Width           =   2655
      Begin VB.TextBox Text44 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   2040
         TabIndex        =   61
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text43 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1080
         TabIndex        =   60
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text42 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   120
         TabIndex        =   59
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text41 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   2040
         TabIndex        =   58
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox Text40 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1080
         TabIndex        =   57
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox Text39 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   120
         TabIndex        =   56
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox Text38 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   2040
         TabIndex        =   55
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text37 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1080
         TabIndex        =   54
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text34 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1080
         TabIndex        =   53
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text35 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   2040
         TabIndex        =   52
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text36 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   120
         TabIndex        =   51
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text33 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   120
         TabIndex        =   50
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "W0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   37
         Top             =   525
         Width           =   375
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "W1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   36
         Top             =   520
         Width           =   255
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "W2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   35
         Top             =   525
         Width           =   375
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFF80&
      Caption         =   "ERROR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   5160
      TabIndex        =   33
      Top             =   3240
      Width           =   975
      Begin VB.TextBox Text32 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   220
         TabIndex        =   49
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text31 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   220
         TabIndex        =   48
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text30 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   220
         TabIndex        =   47
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox Text29 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   220
         TabIndex        =   46
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFF80&
      Caption         =   "OUTPUT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   4200
      TabIndex        =   32
      Top             =   3240
      Width           =   975
      Begin VB.TextBox Text28 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   220
         TabIndex        =   45
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text27 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   220
         TabIndex        =   44
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text26 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   220
         TabIndex        =   43
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox Text25 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   220
         TabIndex        =   42
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF80&
      Caption         =   "HASIL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   8160
      TabIndex        =   31
      Top             =   240
      Width           =   975
      Begin VB.TextBox Text24 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   220
         TabIndex        =   41
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text23 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   220
         TabIndex        =   40
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text22 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   220
         TabIndex        =   39
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox Text21 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   220
         TabIndex        =   38
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF80&
      Caption         =   "THRESHOLD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   28
      Top             =   4440
      Width           =   3615
      Begin VB.TextBox Text20 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   480
         TabIndex        =   30
         Top             =   420
         Width           =   3015
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   450
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "INITIAL WEIGHT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   21
      Top             =   3240
      Width           =   3615
      Begin VB.TextBox Text19 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   3000
         TabIndex        =   27
         Top             =   420
         Width           =   495
      End
      Begin VB.TextBox Text18 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1800
         TabIndex        =   25
         Top             =   420
         Width           =   495
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   600
         TabIndex        =   23
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "W2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   26
         Top             =   450
         Width           =   375
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "W1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   24
         Top             =   450
         Width           =   375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "W0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   450
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "OPERASI AND"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   3000
         TabIndex        =   20
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   3000
         TabIndex        =   19
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   3000
         TabIndex        =   18
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   3000
         TabIndex        =   17
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   2040
         TabIndex        =   16
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1080
         TabIndex        =   15
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   2040
         TabIndex        =   13
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1080
         TabIndex        =   12
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   2040
         TabIndex        =   10
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1080
         TabIndex        =   9
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   2040
         TabIndex        =   6
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3150
         TabIndex        =   7
         Top             =   520
         Width           =   255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "X2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   520
         Width           =   255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "X1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   520
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "X0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   520
         Width           =   255
      End
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Threshold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   72
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "W2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   67
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "W1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   66
      Top             =   1425
      Width           =   375
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "W0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   65
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4530
      TabIndex        =   64
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4530
      TabIndex        =   63
      Top             =   1605
      Width           =   255
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4455
      TabIndex        =   62
      Top             =   750
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   2745
      Left            =   4200
      Picture         =   "JSTand.frx":3538A
      Top             =   240
      Width           =   3750
   End
   Begin VB.Menu fl 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu brd 
         Caption         =   "Beranda"
         Index           =   0
      End
      Begin VB.Menu KLR 
         Caption         =   "Keluar"
         Index           =   0
      End
   End
End
Attribute VB_Name = "JSTand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################# Beranda #################'
Private Sub brd_Click(Index As Integer)
    Beranda.Show
    JSTand.Hide
    Unload Me
End Sub



'####################### Keluar #######################'
Private Sub KLR_Click(Index As Integer)
    Q = MsgBox("Anda yakin akan keluar ?", vbQuestion + vbOKCancel, "System")
    If Q = vbOK Then
        Unload Me
        End
    End If
End Sub



'################# Input Operasi And #################'
'X-Y Baris 1'
Private Sub Text1_Change()
    IO.Xa0 = Val(Text1.Text)
    
    If (IO.Xa0 = Val(0)) Then
        IO.Xa0 = Val(0)
        Text13.Text = Val(IO.Xa0)
    Else
        IO.Xa0 = Val(1)
        Text13.Text = Val(IO.Xa0)
    End If
    
    Text13.Text = Val(IO.Xa0) * Val(IO.Xa1) * Val(IO.Xa2)
    IO.Y0 = Text13.Text
    Text13.Enabled = False
End Sub

Private Sub Text2_Change()
    IO.Xa1 = Val(Text2.Text)
    
    If (IO.Xa1 = Val(0)) Then
        IO.Xa1 = Val(0)
        Text13.Text = Val(IO.Xa1)
    Else
        IO.Xa1 = Val(1)
        Text13.Text = Val(IO.Xa1)
    End If
    
    Text13.Text = Val(IO.Xa0) * Val(IO.Xa1) * Val(IO.Xa2)
    IO.Y0 = Text13.Text
    Text13.Enabled = False
End Sub

Private Sub Text3_Change()
    IO.Xa2 = Val(Text3.Text)
    
    If (IO.Xa2 = Val(0)) Then
        IO.Xa2 = Val(0)
        Text13.Text = Val(IO.Xa2)
    Else
        IO.Xa2 = Val(1)
        Text13.Text = Val(IO.Xa2)
    End If
    
    Text13.Text = Val(IO.Xa0) * Val(IO.Xa1) * Val(IO.Xa2)
    IO.Y0 = Text13.Text
    Text13.Enabled = False
End Sub

'X-Y Baris 2'
Private Sub Text4_Change()
    IO.Xb0 = Val(Text4.Text)
    
    If (IO.Xb0 = Val(0)) Then
        IO.Xb0 = Val(0)
        Text14.Text = Val(IO.Xb0)
    Else
        IO.Xb0 = Val(1)
        Text14.Text = Val(IO.Xb0)
    End If
    
    Text14.Text = Val(IO.Xb0) * Val(IO.Xb1) * Val(IO.Xb2)
    IO.Y1 = Text14.Text
    Text14.Enabled = False
End Sub

Private Sub Text5_Change()
    IO.Xb1 = Val(Text5.Text)
    
    If (IO.Xb1 = Val(0)) Then
        IO.Xb1 = Val(0)
        Text14.Text = Val(IO.Xb1)
    Else
        IO.Xb1 = Val(1)
        Text14.Text = Val(IO.Xb1)
    End If
    
    Text14.Text = Val(IO.Xb0) * Val(IO.Xb1) * Val(IO.Xb2)
    IO.Y1 = Text14.Text
    Text14.Enabled = False
End Sub

Private Sub Text6_Change()
    IO.Xb2 = Val(Text6.Text)
    
    If (IO.Xb2 = Val(0)) Then
        IO.Xb2 = Val(0)
        Text14.Text = Val(IO.Xb2)
    Else
        IO.Xb2 = Val(1)
        Text14.Text = Val(IO.Xb2)
    End If
    
    Text14.Text = Val(IO.Xb0) * Val(IO.Xb1) * Val(IO.Xb2)
    IO.Y1 = Text14.Text
    Text14.Enabled = False
End Sub

'X-Y Baris 3'
Private Sub Text7_Change()
    IO.Xc0 = Val(Text7.Text)
    
    If (IO.Xc0 = Val(0)) Then
        IO.Xc0 = Val(0)
        Text15.Text = Val(IO.Xc0)
    Else
        IO.Xc0 = Val(1)
        Text15.Text = Val(IO.Xc0)
    End If
    
    Text15.Text = Val(IO.Xc0) * Val(IO.Xc1) * Val(IO.Xc2)
    IO.Y2 = Text15.Text
    Text15.Enabled = False
End Sub

Private Sub Text8_Change()
    IO.Xc1 = Val(Text8.Text)
    
    If (IO.Xc1 = Val(0)) Then
        IO.Xc1 = Val(0)
        Text15.Text = Val(IO.Xc1)
    Else
        IO.Xc1 = Val(1)
        Text15.Text = Val(IO.Xc1)
    End If
    
    Text15.Text = Val(IO.Xc0) * Val(IO.Xc1) * Val(IO.Xc2)
    IO.Y2 = Text15.Text
    Text15.Enabled = False
End Sub

Private Sub Text9_Change()
    IO.Xc2 = Val(Text9.Text)
    
    If (IO.Xc2 = Val(0)) Then
        IO.Xc2 = Val(0)
        Text15.Text = Val(IO.Xc2)
    Else
        IO.Xc2 = Val(1)
        Text15.Text = Val(IO.Xc2)
    End If
    
    Text15.Text = Val(IO.Xc0) * Val(IO.Xc1) * Val(IO.Xc2)
    IO.Y2 = Text15.Text
    Text15.Enabled = False
End Sub

'X-Y Baris 4'
Private Sub Text10_Change()
    IO.Xd0 = Val(Text10.Text)
    
    If (IO.Xd0 = Val(0)) Then
        IO.Xd0 = Val(0)
        Text16.Text = Val(IO.Xd0)
    Else
        IO.Xd0 = Val(1)
        Text16.Text = Val(IO.Xd0)
    End If
    
    Text16.Text = Val(IO.Xd0) * Val(IO.Xd1) * Val(IO.Xd2)
    IO.Y3 = Text16.Text
    Text16.Enabled = False
End Sub

Private Sub Text11_Change()
    IO.Xd1 = Val(Text11.Text)
    
    If (IO.Xd1 = Val(0)) Then
        IO.Xd1 = Val(0)
        Text16.Text = Val(IO.Xd1)
    Else
        IO.Xd1 = Val(1)
        Text16.Text = Val(IO.Xd1)
    End If
    
    Text16.Text = Val(IO.Xd0) * Val(IO.Xd1) * Val(IO.Xd2)
    IO.Y3 = Text16.Text
    Text16.Enabled = False
End Sub

Private Sub Text12_Change()
    IO.Xd2 = Val(Text12.Text)
    
    If (IO.Xd2 = Val(0)) Then
        IO.Xd2 = Val(0)
        Text16.Text = Val(IO.Xd2)
    Else
        IO.Xd2 = Val(1)
        Text16.Text = Val(IO.Xd2)
    End If
    
    Text16.Text = Val(IO.Xd0) * Val(IO.Xd1) * Val(IO.Xd2)
    IO.Y3 = Text16.Text
    Text16.Enabled = False
End Sub



'################# Input Bobot & Threshold #################'
Private Sub Text17_Change()
    Text45.Text = Val(Text17.Text)
    IO.W0 = Text45.Text
    Text45.Enabled = False
End Sub

Private Sub Text18_Change()
    Text46.Text = Val(Text18.Text)
    IO.W1 = Text46.Text
    Text46.Enabled = False
End Sub

Private Sub Text19_Change()
    Text47.Text = Val(Text19.Text)
    IO.W2 = Text47.Text
    Text47.Enabled = False
End Sub

Private Sub Text20_Change()
    Text48.Text = Val(Text20.Text)
    IO.T = Text48.Text
    Text48.Enabled = False
End Sub



'####################### Terapkan #######################'
Private Sub Command1_Click()
    'Hasil'
    Text21.Text = (Val(IO.Xa0) * Val(IO.W0)) + (Val(IO.Xa1) * Val(IO.W1)) + (Val(IO.Xa2) * Val(IO.W2))
    Text22.Text = (Val(IO.Xb0) * Val(IO.W0)) + (Val(IO.Xb1) * Val(IO.W1)) + (Val(IO.Xb2) * Val(IO.W2))
    Text23.Text = (Val(IO.Xc0) * Val(IO.W0)) + (Val(IO.Xc1) * Val(IO.W1)) + (Val(IO.Xc2) * Val(IO.W2))
    Text24.Text = (Val(IO.Xd0) * Val(IO.W0)) + (Val(IO.Xd1) * Val(IO.W1)) + (Val(IO.Xd2) * Val(IO.W2))
    
    IO.H0 = Text21.Text
    IO.H1 = Text22.Text
    IO.H2 = Text23.Text
    IO.H3 = Text24.Text
    
    Text21.Enabled = False
    Text22.Enabled = False
    Text23.Enabled = False
    Text24.Enabled = False
    
    'Output'
    If (IO.H0 >= IO.T) Then
        Text25.Text = Val(1)
    ElseIf (IO.H0 < IO.T) Then
        Text25.Text = Val(0)
    End If
    
    If (IO.H1 >= IO.T) Then
        Text26.Text = Val(1)
    ElseIf (IO.H1 < IO.T) Then
        Text26.Text = Val(0)
    End If
    
    If (IO.H2 >= IO.T) Then
        Text27.Text = Val(1)
    ElseIf (IO.H2 < IO.T) Then
        Text27.Text = Val(0)
    End If
    
    If (IO.H3 >= IO.T) Then
        Text28.Text = Val(1)
    ElseIf (IO.H3 < IO.T) Then
        Text28.Text = Val(0)
    End If
    
    IO.Op0 = Text25.Text
    IO.Op1 = Text26.Text
    IO.Op2 = Text27.Text
    IO.Op3 = Text28.Text
    
    Text25.Enabled = False
    Text26.Enabled = False
    Text27.Enabled = False
    Text28.Enabled = False
    
    'Error'
    If (IO.Y0 = IO.Op0) Then
        Text29.Text = Val(0)
    ElseIf (IO.Y0 <> IO.Op0) Then
        Text29.Text = Val(1)
    End If
    
    If (IO.Y1 = IO.Op1) Then
        Text30.Text = Val(0)
    ElseIf (IO.Y1 <> IO.Op1) Then
        Text30.Text = Val(1)
    End If
    
    If (IO.Y2 = IO.Op2) Then
        Text31.Text = Val(0)
    ElseIf (IO.Y2 <> IO.Op2) Then
        Text31.Text = Val(1)
    End If
    
    If (IO.Y3 = IO.Op3) Then
        Text32.Text = Val(0)
    ElseIf (IO.Y3 <> IO.Op3) Then
        Text32.Text = Val(1)
    End If
    
    IO.Err0 = Text29.Text
    IO.Err1 = Text30.Text
    IO.Err2 = Text31.Text
    IO.Err3 = Text32.Text
    
    Text29.Enabled = False
    Text30.Enabled = False
    Text31.Enabled = False
    Text32.Enabled = False
    
    'Updated Weight (Bias)'
    Text33.Text = IO.W0
    Text34.Text = IO.W1
    Text35.Text = IO.W2
    Text36.Text = IO.W0
    Text37.Text = IO.W1
    Text38.Text = IO.W2
    Text39.Text = IO.W0
    Text40.Text = IO.W1
    Text41.Text = IO.W2
    Text42.Text = IO.W0
    Text43.Text = IO.W1
    Text44.Text = IO.W2
    
    IO.UWa0 = Text33.Text
    IO.UWa1 = Text34.Text
    IO.UWa2 = Text35.Text
    IO.UWb0 = Text36.Text
    IO.UWb1 = Text37.Text
    IO.UWb2 = Text38.Text
    IO.UWc0 = Text39.Text
    IO.UWc1 = Text40.Text
    IO.UWc2 = Text41.Text
    IO.UWd0 = Text42.Text
    IO.UWd1 = Text43.Text
    IO.UWd2 = Text44.Text
    
    Text33.Enabled = False
    Text34.Enabled = False
    Text35.Enabled = False
    Text36.Enabled = False
    Text37.Enabled = False
    Text38.Enabled = False
    Text39.Enabled = False
    Text40.Enabled = False
    Text41.Enabled = False
    Text42.Enabled = False
    Text43.Enabled = False
    Text44.Enabled = False
    MsgBox "Data Berhasil Diproses, tekan ok", vbInformation, "Notifikasi JST-AND"
End Sub



'####################### Hapus #######################'
Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    Text19.Text = ""
    Text20.Text = ""
    Text21.Text = ""
    Text22.Text = ""
    Text23.Text = ""
    Text24.Text = ""
    Text25.Text = ""
    Text26.Text = ""
    Text27.Text = ""
    Text28.Text = ""
    Text29.Text = ""
    Text30.Text = ""
    Text31.Text = ""
    Text32.Text = ""
    Text33.Text = ""
    Text34.Text = ""
    Text35.Text = ""
    Text36.Text = ""
    Text37.Text = ""
    Text38.Text = ""
    Text39.Text = ""
    Text40.Text = ""
    Text41.Text = ""
    Text42.Text = ""
    Text43.Text = ""
    Text44.Text = ""
    Text45.Text = ""
    Text46.Text = ""
    Text47.Text = ""
    Text48.Text = ""
    MsgBox "Data Berhasil Dibersihkan, tekan ok", vbInformation, "Notifikasi JST-AND"
    Text1.SetFocus
End Sub
