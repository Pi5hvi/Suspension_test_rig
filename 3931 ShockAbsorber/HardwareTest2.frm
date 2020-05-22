VERSION 5.00
Begin VB.Form HardwareTest2 
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
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
            TabIndex        =   6
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
               Left            =   6600
               TabIndex        =   9
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
               ItemData        =   "HardwareTest2.frx":0000
               Left            =   5280
               List            =   "HardwareTest2.frx":003D
               Style           =   2  'Dropdown List
               TabIndex        =   8
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
               Left            =   3240
               TabIndex        =   7
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
               Left            =   4560
               TabIndex        =   11
               Top             =   5640
               Width           =   1575
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H80000008&
               Height          =   1920
               Left            =   1800
               ScaleHeight     =   124
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   496
               TabIndex        =   10
               Top             =   3120
               Width           =   7500
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H80000008&
               Height          =   1920
               Left            =   1800
               ScaleHeight     =   124
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   496
               TabIndex        =   4
               Top             =   840
               Width           =   7500
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
               Left            =   1800
               TabIndex        =   5
               Top             =   120
               Width           =   7455
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
Attribute VB_Name = "HardwareTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim currRpm As Integer
Dim valueArr(20, 1, 100) As Integer

Private Sub Command1_Click()
End Sub

Private Sub Command2_Click()
End Sub

Private Sub Command3_Click()
    Unload Me
    Form3.Show
End Sub



Private Sub Command4_Click()
Dim yyCurr As Integer, yyCurr1 As Integer, myx As Integer, xx As Integer, yyPrev As Integer, yyPrev1 As Integer
currRpm = Combo1.ListIndex
For xx = 0 To 99
    yyCurr = (255 - valueArr(currRpm, 0, xx)) / 2
    yyCurr1 = (255 - valueArr(currRpm, 1, xx)) / 2
    myx = xx * 5
    If xx <> 0 Then
        Picture1.Line (myx - 5, yyPrev)-(myx, yyCurr), RGB(255, 0, 0)
        Picture2.Line (myx - 5, yyPrev1)-(myx, yyCurr1), RGB(0, 255, 0)
    End If

    yyPrev = yyCurr
    yyPrev1 = yyCurr1
    'xx = xx + 1
Next
End Sub

Private Sub Form_Load()
    Combo1.ListIndex = 0
    readFile
End Sub
Sub readFile()
Dim i As Integer, j As Integer, k As Integer, fnum As Integer
fnum = FreeFile
Open App.Path & "\myfile.txt" For Input As fnum
For i = 0 To 19
    For j = 0 To 1
        For k = 0 To 99
            Input #fnum, valueArr(i, j, k)
        Next
    Next
Next
Close fnum
End Sub


