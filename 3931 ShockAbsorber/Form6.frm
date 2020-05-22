VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form6 
   BackColor       =   &H000080FF&
   Caption         =   "Form6"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "Form6"
   ScaleHeight     =   4335
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "BACK"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   3840
      Width           =   8055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Stepper Motor"
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   6120
      TabIndex        =   16
      Top             =   1200
      Width           =   2055
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Form6.frx":0000
         Left            =   120
         List            =   "Form6.frx":000D
         TabIndex        =   21
         Text            =   "Combo1"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Text            =   "0"
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         Height          =   855
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1815
         Begin VB.OptionButton Option2 
            BackColor       =   &H00000000&
            Caption         =   "Anticlockwise"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00000000&
            Caption         =   "Clockwise"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   7560
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "DC Motor Control"
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   3120
      TabIndex        =   8
      Top             =   1200
      Width           =   2895
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form6.frx":001A
         Left            =   1440
         List            =   "Form6.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form6.frx":001E
         Left            =   1440
         List            =   "Form6.frx":0020
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Send Command"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Off Time"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Port"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "On Time"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "MOTOR 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      Caption         =   "DC Motor Control"
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2895
      Begin VB.CommandButton Command2 
         Caption         =   "Send Command"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   2535
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "Form6.frx":0022
         Left            =   1440
         List            =   "Form6.frx":0024
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         ItemData        =   "Form6.frx":0026
         Left            =   1440
         List            =   "Form6.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "MOTOR 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "On Time"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Port"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label22 
         BackColor       =   &H00000000&
         Caption         =   "Off Time"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H000080FF&
      Caption         =   "INDUSTRIAL MONITORING AND CONTROL SYSTEM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   7935
   End
   Begin VB.Label Label14 
      BackColor       =   &H000080FF&
      Caption         =   "HARDWARE TEST"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2760
      TabIndex        =   22
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MSComm1.Output = Chr(75)  ' command
    MSComm1.Output = Chr(0)  'port1 Pin
    MSComm1.Output = Chr(Combo1.ListIndex) 'data1
    MSComm1.Output = Chr(Combo2.ListIndex) 'data1
End Sub

Private Sub Command2_Click()
    MSComm1.Output = Chr(75)  ' command
    MSComm1.Output = Chr(1)  'port1 Pin
    MSComm1.Output = Chr(Combo6.ListIndex) 'data1
    MSComm1.Output = Chr(Combo7.ListIndex) 'data1
End Sub

Private Sub Command3_Click()
Unload Me
    Form3.Show

End Sub

Private Sub Form_Load()
    Dim i As Integer
    MSComm1.PortOpen = True
    For i = 0 To 255
        Combo6.AddItem i
        Combo7.AddItem i
        Combo1.AddItem i
        Combo2.AddItem i
    Next
    Combo3.ListIndex = 0
    
    MSComm1.Output = Chr(28)
    MSComm1.Output = Chr(0)
    MSComm1.Output = Chr(29)
    MSComm1.Output = Chr(0)
    MSComm1.Output = Chr(30)
    MSComm1.Output = Chr(0)
    MSComm1.Output = Chr(31)
    MSComm1.Output = Chr(0)
    MSComm1.Output = Chr(34)
    MSComm1.Output = Chr(0)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
MSComm1.PortOpen = False
End Sub

Private Sub Text2_Change()
    MSComm1.Output = Chr(74)  ' command
    MSComm1.Output = Chr(3)  'port
    If Option1.Value = True Then
        MSComm1.Output = Chr(0)  '
    Else
        MSComm1.Output = Chr(1)  'clk / anti
    End If
    MSComm1.Output = Chr(CInt(Val(Combo3.ListIndex)))  ' mode
    If Val(Text2) <= 255 Then
        MSComm1.Output = Chr(CByte(Val(Text2)))  'speed
    End If
End Sub
