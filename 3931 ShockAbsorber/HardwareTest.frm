VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   1200
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9240
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   9840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      Begin VB.Frame Frame2 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   7335
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   11655
         Begin VB.Frame Frame4 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   11175
            Begin VB.CommandButton Command4 
               Caption         =   "SET"
               BeginProperty Font 
                  Name            =   "Maiandra GD"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   3480
               TabIndex        =   15
               Top             =   120
               Width           =   1095
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               ItemData        =   "HardwareTest.frx":0000
               Left            =   2160
               List            =   "HardwareTest.frx":003D
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               Caption         =   "SELECT RPM : "
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   13
               Top             =   120
               Width           =   1935
            End
         End
         Begin VB.Timer TimerRefresh 
            Enabled         =   0   'False
            Interval        =   150
            Left            =   2880
            Top             =   240
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   6255
            Left            =   120
            TabIndex        =   3
            Top             =   840
            Width           =   11175
            Begin VB.CommandButton Command1 
               Caption         =   "START SCAN"
               Enabled         =   0   'False
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
               Left            =   120
               TabIndex        =   19
               Top             =   5640
               Width           =   2175
            End
            Begin VB.CommandButton Command2 
               Caption         =   "STOP SCAN"
               Enabled         =   0   'False
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
               Left            =   2400
               TabIndex        =   18
               Top             =   5640
               Width           =   2175
            End
            Begin VB.CommandButton Command3 
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
               Left            =   9480
               TabIndex        =   17
               Top             =   5640
               Width           =   1575
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H80000008&
               Height          =   1920
               Left            =   3480
               ScaleHeight     =   124
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   496
               TabIndex        =   16
               Top             =   3120
               Width           =   7500
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H80000008&
               Height          =   1920
               Left            =   3480
               ScaleHeight     =   124
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   496
               TabIndex        =   6
               Top             =   840
               Width           =   7500
            End
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   4185
               Index           =   0
               Left            =   120
               TabIndex        =   4
               Top             =   480
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   7382
               _Version        =   393216
               Appearance      =   1
               Max             =   255
               Orientation     =   1
               Scrolling       =   1
            End
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   4185
               Index           =   1
               Left            =   1680
               TabIndex        =   5
               Top             =   480
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   7382
               _Version        =   393216
               Appearance      =   1
               Max             =   255
               Orientation     =   1
               Scrolling       =   1
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               Caption         =   "CHANNEL 0"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   11
               Top             =   4800
               Width           =   1455
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               Caption         =   "CHANNEL 0"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   10
               Top             =   4800
               Width           =   1455
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               Caption         =   "GRAPH"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   3480
               TabIndex        =   9
               Top             =   120
               Width           =   7455
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               Caption         =   "F0"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1680
               TabIndex        =   8
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               Caption         =   "FT"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   7
               Top             =   120
               Width           =   1455
            End
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "TITLE CONTROL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   10815
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adc(2) As Integer, currentChannel As Integer, myCount As Integer
Dim deviceOP As Integer
Dim yyCurr As Integer, yyCurr1 As Integer, xx As Integer, yyPrev As Integer, yyPrev1 As Integer
Dim flag As Boolean, currRpm As Integer
Dim valueArr(20, 1, 100) As Integer

Private Sub Command1_Click()
    Timer1.Enabled = True
    TimerRefresh.Enabled = True
End Sub

Private Sub Command2_Click()
storeFile
Timer1.Enabled = False
TimerRefresh.Enabled = False
End Sub

Private Sub Command3_Click()
    Unload Me
    Form3.Show
End Sub



Private Sub Command4_Click()
currRpm = Combo1.ListIndex
Command1.Enabled = True
Command2.Enabled = True
End Sub

Private Sub Form_Load()
    currentChannel = 0
    MSComm1.CommPort = port
    MSComm1.PortOpen = True
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
MSComm1.Output = Chr(34)
MSComm1.Output = Chr(0)
End Sub

Private Sub MSComm1_OnComm()
Dim temp As Integer
If MSComm1.CommEvent = comEvReceive Then
    temp = Asc(MSComm1.Input)
    myCount = myCount + 1
    If myCount = 10 Then
       myCount = 0
        adc(currentChannel) = temp
        currentChannel = currentChannel + 1
        If currentChannel >= 2 Then
            currentChannel = 0
        End If
     End If
End If
End Sub
Private Sub Timer1_Timer()
    MSComm1.Output = Chr(73)
    MSComm1.Output = Chr(currentChannel)
    Debug.Print "CR CHANNEL :" & currentChannel
End Sub

Private Sub TimerRefresh_Timer()
Dim i As Integer, maxHeight As Integer, myx As Integer
    For i = 0 To 1
        ProgressBar1(i).Value = adc(i)
        Label5(i).Caption = adc(i)
        valueArr(currRpm, i, xx) = adc(i)
    Next
'   drawGraph
    yyCurr = (255 - adc(0)) / 2
    yyCurr1 = (255 - adc(1)) / 2
    myx = xx * 5
    If xx <> 0 Then
        Picture1.Line (myx - 5, yyPrev)-(myx, yyCurr), RGB(255, 0, 0)
        Picture2.Line (myx - 5, yyPrev1)-(myx, yyCurr1), RGB(0, 255, 0)
    End If

    yyPrev = yyCurr
    yyPrev1 = yyCurr1
    xx = xx + 1
    If xx = 100 Then
        xx = 0
        
        Timer1.Enabled = False
        TimerRefresh.Enabled = False
        Command1.Enabled = False
        Command2.Enabled = False
        storeFile
        'Picture1.Cls
    End If

End Sub
Private Sub drawGraph()

End Sub

Sub storeFile()
Dim i As Integer, j As Integer, k As Integer, fnum As Integer
fnum = FreeFile
Open App.Path & "\myfile.txt" For Output As fnum
For i = 0 To 19
    For j = 0 To 1
        For k = 0 To 99
            Print #fnum, valueArr(i, j, k)
        Next
    Next
Next
Close fnum
End Sub

