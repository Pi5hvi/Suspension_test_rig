VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PlotGraph 
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   5760
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE TO EXCEL"
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
      Left            =   3720
      TabIndex        =   11
      Top             =   6480
      Width           =   4215
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
            TabIndex        =   6
            Top             =   120
            Width           =   11175
            Begin VB.CommandButton Command4 
               Caption         =   "PLOT"
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
               Left            =   3360
               TabIndex        =   7
               Top             =   120
               Width           =   4215
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
            Begin VB.OptionButton Option2 
               BackColor       =   &H000080FF&
               Caption         =   "LINE GRAPH"
               Height          =   375
               Left            =   4560
               TabIndex        =   10
               Top             =   720
               Width           =   2415
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H000080FF&
               Caption         =   "BAR GRAPH"
               Height          =   375
               Left            =   1800
               TabIndex        =   9
               Top             =   720
               Value           =   -1  'True
               Width           =   2655
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
               Left            =   4680
               TabIndex        =   8
               Top             =   5640
               Width           =   1575
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H80000008&
               FillColor       =   &H00FFFFFF&
               FillStyle       =   0  'Solid
               Height          =   3600
               Left            =   1800
               ScaleHeight     =   236
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   496
               TabIndex        =   4
               Top             =   1200
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
         Caption         =   "TILT CONTROL"
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
Attribute VB_Name = "PlotGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim currRpm As Integer
Dim valueArr(20, 1, 100) As Integer
Dim calcArr(20) As Double


Private Sub Command1_Click()
Dim fnum As Integer, fname As String, i As Integer, sstr As String
cd.CancelError = False
cd.InitDir = App.Path
cd.ShowSave
If cd.FileName <> "" Then
    fnum = FreeFile
    fname = cd.FileName
    fname = fname & ".csv"
    Open fname For Output As fnum
        sstr = "Sr No. ,Tr,N(RPM)"
        Print #fnum, sstr
        For i = 0 To 19
            sstr = (i + 1) & "," & calcArr(i) & "," & ((i * 5) + 5)
            Print #fnum, sstr
        Next
    Close fnum
End If

End Sub

Private Sub Command2_Click()
End Sub

Private Sub Command3_Click()
    Unload Me
    Form3.Show
End Sub



Private Sub Command4_Click()
Dim i As Integer, myVal As Double, j As Integer, myVal2 As Double, k As Integer
Dim yScale As Double, yVal As Double, xx As Double, yy As Double, xx2 As Double, yy2 As Double
Picture1.Cls
Picture1.Line (98, 0)-(98, 240), RGB(255, 255, 0)
Picture1.Line (0, 220)-(500, 220), RGB(255, 255, 0)
yScale = 74
Picture1.ForeColor = vbRed
Picture1.Font.Size = 7

Picture1.CurrentX = 25
Picture1.CurrentY = 222
Picture1.Print "RPM"

Picture1.CurrentX = 75
Picture1.CurrentY = 0
Picture1.Print "3.0-"

Picture1.CurrentX = 75
Picture1.CurrentY = 20
Picture1.Print "2.7-"

Picture1.CurrentX = 75
Picture1.CurrentY = 40
Picture1.Print "2.4-"

Picture1.CurrentX = 75
Picture1.CurrentY = 60
Picture1.Print "2.1-"

Picture1.CurrentX = 75
Picture1.CurrentY = 80
Picture1.Print "1.8-"

Picture1.CurrentX = 75
Picture1.CurrentY = 100
Picture1.Print "1.5-"

Picture1.CurrentX = 75
Picture1.CurrentY = 120
Picture1.Print "1.2-"

Picture1.CurrentX = 75
Picture1.CurrentY = 140
Picture1.Print "0.9-"

Picture1.CurrentX = 75
Picture1.CurrentY = 160
Picture1.Print "0.6-"

Picture1.CurrentX = 75
Picture1.CurrentY = 180
Picture1.Print "0.3-"

Picture1.CurrentX = 75
Picture1.CurrentY = 200
Picture1.Print "0.0-"

Picture1.CurrentY = 235
For i = 0 To 19
    myVal = 0
    myVal2 = 0
    For j = 0 To 99
        myVal = myVal + valueArr(i, 0, j)
        myVal2 = myVal2 + valueArr(i, 1, j)
    Next
    myVal = myVal / 100#
    myVal2 = myVal2 / 100#
    If myVal = 0 Then
        calcArr(i) = 0
        yy = 220
    Else
        calcArr(i) = myVal / myVal2
        yVal = calcArr(i) * yScale
        yy = 220 - yVal
    End If
    xx = ((i + 1) * 20) + 100
    xx2 = ((i) * 20) + 100
    Debug.Print calcArr(i)
    If Option1.Value = True Then
        Picture1.Line (xx, yy)-(xx + 15, 220), 255, BF
    Else
        'If (i > 0) Then
            Picture1.Line (xx2, yy2)-(xx, yy), 255
        'End If
    End If
    yy2 = yy
    Picture1.CurrentX = xx2 '- 10
    Picture1.CurrentY = 222
    Picture1.Print ((i * 5) + 10)
Next
End Sub

Private Sub Form_Load()
    Dim i As Integer, dbl As Double
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


