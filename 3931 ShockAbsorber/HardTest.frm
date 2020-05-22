VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000008&
   Caption         =   "Form2"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7005
   LinkTopic       =   "Form2"
   ScaleHeight     =   3885
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command14 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Sensor's"
         ForeColor       =   &H8000000E&
         Height          =   1215
         Left            =   360
         TabIndex        =   5
         Top             =   1680
         Width           =   6255
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   375
            Index           =   0
            Left            =   2160
            TabIndex        =   6
            Top             =   120
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Max             =   255
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   375
            Index           =   1
            Left            =   2160
            TabIndex        =   7
            Top             =   600
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Max             =   255
         End
         Begin VB.Label Label2 
            BackColor       =   &H00808080&
            Caption         =   "FT"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   480
            TabIndex        =   11
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackColor       =   &H00808080&
            Caption         =   "F0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   480
            TabIndex        =   10
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label10 
            BackColor       =   &H00808080&
            Caption         =   "Label10"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   4920
            TabIndex        =   9
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label10 
            BackColor       =   &H00808080&
            Caption         =   "Label10"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   4920
            TabIndex        =   8
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "START"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Timer TimerRefresh 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1320
         Top             =   3600
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   720
         Top             =   3600
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   360
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         CommPort        =   3
         DTREnable       =   -1  'True
         RThreshold      =   1
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   6375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "SHOCK ABSOBER"
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
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   6495
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adc(2) As Integer, currentChannel As Integer, myCount As Integer
Dim deviceOP As Integer, temp As Integer
Dim flag As Boolean

Private Sub Command1_Click()
Timer1.Enabled = True
TimerRefresh.Enabled = True
End Sub

Private Sub Command14_Click()
Unload Me
Form3.Show
End Sub

Private Sub Form_Load()
MSComm1.CommPort = Module1.port
MSComm1.PortOpen = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
TimerRefresh.Enabled = False

End Sub

Private Sub MSComm1_OnComm()
If MSComm1.CommEvent = comEvReceive Then
    temp = Asc(MSComm1.Input)
    myCount = myCount + 1
    If myCount = 10 Then
        myCount = 0
        adc(currentChannel) = temp
        currentChannel = currentChannel + 1
        If currentChannel = 2 Then
            currentChannel = 0
        End If
     End If
End If
End Sub

Private Sub Timer1_Timer()
    MSComm1.Output = Chr(73)
    MSComm1.Output = Chr(currentChannel)
    Debug.Print "CURRENT CHANNEL:" & currentChannel
End Sub


Private Sub TimerRefresh_Timer()
Dim i As Integer
For i = 0 To 1
        ProgressBar1(i).Value = adc(i)
        Label10(i).Caption = adc(i)
Next
End Sub
