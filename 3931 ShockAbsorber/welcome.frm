VERSION 5.00
Object = "{6F2A8BEC-0B0A-4991-A21A-ED31E9C8005D}#28.0#0"; "mybutton.ocx"
Begin VB.Form Form6 
   BackColor       =   &H000080FF&
   Caption         =   "Form6"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form6"
   ScaleHeight     =   3495
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin vbREGs.CommandButton CommandButton1 
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   2.45745e5
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   ""
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   6975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROCEED"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   6975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.Frame Frame2 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   6495
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   4080
            TabIndex        =   5
            Text            =   "Combo1"
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "SELECT COM PORT :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   3855
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "WELCOME"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   6375
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "SHOCK ABSORBER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   735
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   6495
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Module1.port = Combo1.ListIndex + 1
    Debug.Print Module1.port
    Unload Me
    Form3.Show
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()
    For i = 1 To 15
        Combo1.AddItem ("COM " & i)
    Next
   Combo1.ListIndex = 0
End Sub

