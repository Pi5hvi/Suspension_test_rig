VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   "Form3"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5445
   LinkTopic       =   "Form3"
   ScaleHeight     =   7260
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "FT / F0 GRAPH"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   8
      Top             =   4440
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   4935
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   5415
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   4455
         Begin VB.CommandButton Command7 
            Caption         =   "ABOUT"
            BeginProperty Font 
               Name            =   "Maiandra GD"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   9
            Top             =   4200
            Width           =   3975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "HARDWARE TEST"
            BeginProperty Font 
               Name            =   "Maiandra GD"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   240
            TabIndex        =   7
            Top             =   120
            Width           =   3975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "AUTOMATIC TITLE CONTROL"
            BeginProperty Font 
               Name            =   "Maiandra GD"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   240
            TabIndex        =   6
            Top             =   1320
            Width           =   3975
         End
         Begin VB.CommandButton Command3 
            Caption         =   "VIEW GRAPH"
            BeginProperty Font 
               Name            =   "Maiandra GD"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   240
            TabIndex        =   5
            Top             =   2400
            Width           =   3975
         End
         Begin VB.CommandButton Command4 
            Caption         =   "HELP"
            BeginProperty Font 
               Name            =   "Maiandra GD"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   4
            Top             =   4800
            Width           =   3975
         End
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   6720
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "SHOCK ABSORBER"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
    Form2.Show
End Sub

Private Sub Command2_Click()
     Unload Me
     Form1.Show
End Sub

Private Sub Command3_Click()
    Unload Me
    HardwareTest2.Show
End Sub

Private Sub Command4_Click()
    Unload Me
    Form5.Show
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
    Unload Me
    PlotGraph.Show
End Sub

Private Sub Command7_Click()
    Unload Me
    Form4.Show
End Sub
