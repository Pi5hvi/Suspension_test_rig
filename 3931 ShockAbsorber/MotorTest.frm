VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FormMotor 
   BackColor       =   &H00404040&
   Caption         =   "Form6"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5730
   LinkTopic       =   "Form6"
   ScaleHeight     =   6120
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   2520
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BACK"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   5175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Stepper Motor"
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   5175
      Begin VB.CommandButton Command6 
         Caption         =   "REVERSE CONTINOUS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   9
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "STOP CONTINOUS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   8
         Top             =   2160
         Width           =   4095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "FORWORD  CONTINOUS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   7
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         Height          =   855
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   4215
         Begin VB.CommandButton Command2 
            Caption         =   ">>"
            Height          =   495
            Left            =   2280
            TabIndex        =   6
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "<<"
            Height          =   495
            Left            =   360
            TabIndex        =   5
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5520
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "TILT CONTROL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "MOTOR TEST"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   5175
   End
End
Attribute VB_Name = "FormMotor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim inArr(8) As Integer, cnt As Integer, isReverse As Boolean


Private Sub Command1_Click()
    
    cnt = cnt - 1
    If cnt = -1 Then
        cnt = 7
    End If
    MSComm1.Output = Chr(35)
    MSComm1.Output = Chr(inArr(cnt))

    
    Debug.Print cnt
End Sub

Private Sub Command2_Click()
    
    cnt = cnt + 1
     If cnt = 8 Then
        cnt = 0
    End If
    MSComm1.Output = Chr(35)
    MSComm1.Output = Chr(inArr(cnt))
    
   
 Debug.Print cnt
End Sub

Private Sub Command3_Click()
Unload Me
    Form3.Show

End Sub

Private Sub Command4_Click()
    isReverse = False
    Timer1.Enabled = True
End Sub

Private Sub Command5_Click()
Timer1.Enabled = False
End Sub

Private Sub Command6_Click()
    isReverse = True
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    Dim i As Integer
    MSComm1.CommPort = port
    MSComm1.PortOpen = True
   
    
    inArr(0) = 1
    inArr(1) = 3
    inArr(2) = 2
    inArr(3) = 6
    inArr(4) = 4
    inArr(5) = 12
    inArr(6) = 8
    inArr(7) = 9
    cnt = 0
    
    
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

Private Sub Timer1_Timer()
If isReverse Then
    Command1_Click
Else
    Command2_Click
End If
End Sub
