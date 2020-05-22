VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6F2A8BEC-0B0A-4991-A21A-ED31E9C8005D}#28.0#0"; "mybutton.ocx"
Begin VB.Form FrmDeviceControl 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   11655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14010
   LinkTopic       =   "Form1"
   ScaleHeight     =   11655
   ScaleWidth      =   14010
   StartUpPosition =   2  'CenterScreen
   Begin vbREGs.CommandButton CommandButton3 
      Height          =   495
      Left            =   9000
      TabIndex        =   38
      Top             =   10800
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   873
      Caption         =   "BACK"
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   13920
      Top             =   5760
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   12600
      TabIndex        =   15
      Top             =   7920
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   14760
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   8880
      TabIndex        =   14
      Top             =   7920
      Width           =   3735
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   14160
      Top             =   7800
   End
   Begin VB.PictureBox CommandButton1 
      Height          =   255
      Left            =   8880
      ScaleHeight     =   195
      ScaleWidth      =   2355
      TabIndex        =   5
      Top             =   7320
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      Height          =   5775
      Left            =   8880
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Height          =   6975
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   8415
      Begin MSComctlLib.Slider Slider1 
         Height          =   2655
         Left            =   360
         TabIndex        =   36
         Top             =   4080
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   4683
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   1
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         DrawWidth       =   2
         Height          =   3780
         Left            =   360
         ScaleHeight     =   256.719
         ScaleMode       =   0  'User
         ScaleWidth      =   531
         TabIndex        =   24
         Top             =   120
         Width           =   8025
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "0-"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "0-"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "0-"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "0-"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "0-"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   31
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "0-"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   30
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "0-"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   29
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "0-"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   28
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "0-"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "0-"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "0-"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000007&
         Caption         =   "CHANNEL 8:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   6600
         Width           =   1335
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   2520
         Top             =   6600
         Width           =   3825
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000007&
         Caption         =   "CHANNEL 7:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1080
         TabIndex        =   22
         Top             =   6240
         Width           =   1335
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00004080&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   2520
         Top             =   6240
         Width           =   3825
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000007&
         Caption         =   "CHANNEL6:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   5880
         Width           =   1335
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   2520
         Top             =   5880
         Width           =   3825
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000007&
         Caption         =   "CHANNEL 5:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1080
         TabIndex        =   20
         Top             =   5520
         Width           =   1335
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00008080&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   2520
         Top             =   5520
         Width           =   3825
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000007&
         Caption         =   "CHANNEL 4:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   19
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   2520
         Top             =   5160
         Width           =   3825
      End
      Begin VB.Label deviceLab 
         BackColor       =   &H80000007&
         Caption         =   "CHANNEL 3:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   2520
         Top             =   4440
         Width           =   3825
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000007&
         Caption         =   "CHANNEL 2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   2520
         Top             =   4800
         Width           =   3825
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000007&
         Caption         =   "CHANNEL 1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1080
         TabIndex        =   16
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   2520
         Top             =   4080
         Width           =   3825
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   7680
      Width           =   8415
      Begin vbREGs.CommandButton CommandButton4 
         Height          =   615
         Left            =   4560
         TabIndex        =   37
         Top             =   2640
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1085
         Caption         =   "ADD"
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   6000
         TabIndex        =   13
         Top             =   2040
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   20381699
         CurrentDate     =   40453
      End
      Begin VB.OptionButton Option3 
         Caption         =   "OFF"
         Height          =   495
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1440
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "ON"
         Height          =   495
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3030
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label LblDisp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AT TIME"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   19
         Left            =   4560
         TabIndex        =   12
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label LblDisp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   18
         Left            =   4560
         TabIndex        =   11
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label LblDisp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SELECT DEVICE"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   16
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.PictureBox CommandButton2 
      Height          =   255
      Left            =   11400
      ScaleHeight     =   195
      ScaleWidth      =   2235
      TabIndex        =   6
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   3615
      Left            =   8760
      Top             =   7800
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "NAME DEVICES"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   8880
      TabIndex        =   8
      Top             =   960
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "WIRELESS INDUSTRIAL MONITORTING AND CONTROL SYSTEM"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   13575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   6855
      Left            =   8760
      Top             =   840
      Width           =   5055
   End
End
Attribute VB_Name = "FrmDeviceControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim strFname(8) As String, deviceStatus(8) As Integer, X As Integer
Dim outFlag As Boolean
Dim deviceNo() As Integer, Status() As Integer, atTime() As Date, deviceCount As Integer
Dim ValueArray(7, 300) As Integer
Dim MyColor(7) As Long, currentChannel As Integer
Dim adc(7) As Integer, flag As Boolean, prev As Integer

Private Sub CommandButton3_Click()
    Unload Me
    Form3.Show
End Sub

Private Sub CommandButton4_Click()
Dim s As String
    ReDim Preserve deviceNo(deviceCount)
    ReDim Preserve Status(deviceCount)
    ReDim Preserve atTime(deviceCount)
    deviceNo(deviceCount) = List1.ListIndex
    If Option2.Value = True Then
        Status(deviceCount) = 1
    Else
        Status(deviceCount) = 0
    End If
    atTime(deviceCount) = CDate(DTPicker1.Hour & ":" & DTPicker1.Minute & ":" & DTPicker1.Second)
    s = List1.List(List1.ListIndex) & " "
    If Option2.Value = True Then
        s = s & "ON AT " & DTPicker1.Hour & ":" & DTPicker1.Minute & ":" & DTPicker1.Second
    Else
        s = s & "OFF AT " & DTPicker1.Hour & ":" & DTPicker1.Minute & ":" & DTPicker1.Second
    End If
    List2.AddItem s
    deviceCount = deviceCount + 1
End Sub

Private Sub Form_Load()
Dim i As Integer, myx As Integer, temp As Integer, factor As Integer
MyColor(0) = RGB(0, 0, 255)
MyColor(1) = RGB(255, 0, 0)
MyColor(2) = RGB(255, 80, 80)
Shape4.BackColor = RGB(255, 80, 80)

MyColor(3) = RGB(170, 170, 255)
Shape6.BackColor = RGB(170, 170, 255)
MyColor(4) = RGB(255, 100, 255)
Shape7.BackColor = RGB(255, 100, 255)
MyColor(5) = RGB(100, 255, 255)
Shape8.BackColor = RGB(100, 255, 255)
MyColor(6) = RGB(170, 170, 255)
Shape9.BackColor = RGB(170, 170, 255)
MyColor(7) = RGB(255, 170, 170)
Shape10.BackColor = RGB(255, 170, 170)
MSComm1.CommPort = 3

MSComm1.PortOpen = True
For i = 0 To 7
    deviceStatus(i) = 0
    List1.AddItem "Device " & (i + 1)
Next
factor = 1#
temp = 255# / 10
For myx = 1 To 10
    Label5(myx).Caption = Round((temp * myx) * factor, 1)
Next myx
MSComm1.Output = Chr(34)
    MSComm1.Output = Chr(0)
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MSComm1.PortOpen = False
End Sub

Private Sub MSComm1_OnComm()
Dim fl As Boolean, temp As Integer, ReadVal As Integer, mycnt As Integer, maxWidth As Integer
If MSComm1.CommEvent = comEvReceive Then
    ReadVal = Asc(MSComm1.Input)
        adc(currentChannel) = ReadVal
        If currentChannel = 0 Then
            Shape3.Width = (adc(0) * 15)
        ElseIf currentChannel = 1 Then
            Shape5.Width = (adc(1) * 15)
        ElseIf currentChannel = 2 Then
            Shape4.Width = (adc(2) * 15)
        ElseIf currentChannel = 3 Then
            Shape6.Width = (adc(3) * 15)
        ElseIf currentChannel = 4 Then
            Shape7.Width = (adc(4) * 15)
        ElseIf currentChannel = 5 Then
            Shape8.Width = (adc(5) * 15)
        ElseIf currentChannel = 6 Then
            Shape9.Width = (adc(6) * 15)
        ElseIf currentChannel = 7 Then
            Shape10.Width = (adc(7) * 15)
        End If
        currentChannel = currentChannel + 1
            If currentChannel = 8 Then
                currentChannel = 0
            End If
    End If

End Sub

Private Sub Slider1_Click()
Timer2.Interval = Slider1.Value * 10

End Sub

Private Sub Timer1_Timer()
Dim i As Integer, diff As Long, tempDate As Date
Dim hh As Integer, mm As Integer, ss As Integer, temp As Integer, outData As Integer
Label2.Caption = "Time :-   " & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
    tempDate = CDate(Hour(Now) & ":" & Minute(Now) & ":" & Second(Now))
    List3.Clear
    For i = 0 To deviceCount - 1
        diff = DateDiff("s", tempDate, atTime(i))
    If diff < 0 Then
            List3.AddItem "DONE"
        Else
            mm = diff \ 60
            hh = mm \ 60
            mm = mm Mod 60
            ss = diff Mod 60
            List3.AddItem hh & ":" & mm & ":" & ss
            If diff = 0 Then
                deviceStatus(deviceNo(i)) = Status(i)
        End If
    End If
    Next
    
    temp = 1
    outData = 0
    For i = 0 To 7
        If deviceStatus(i) = 1 Then outData = outData + temp
        temp = temp * 2
    Next
        If flag = True Then
            If prev <> outData Then
                MSComm1.Output = Chr(34)
                MSComm1.Output = Chr(outData)
                
            End If
            prev = outData
            flag = False
        Else
            MSComm1.Output = Chr(73)
            MSComm1.Output = Chr(currentChannel)
            flag = True
        End If
      'PlotGraph
End Sub
Private Sub PlotGraph()
Dim temp As Integer
For temp = 0 To 7
    ValueArray(temp, (X \ 3) + 1) = 260 - adc(temp)
Next temp
For temp = 0 To 7
    
        Picture1.Line (X, ValueArray(temp, (X \ 3)))-(X + 3, ValueArray(temp, (X \ 3) + 1)), MyColor(temp)
    
Next temp
temp = currentChannel - 1
X = X + 3
        If X = 534 Then
            X = 0
            Picture1.Cls
            Picture1.Line (0, 60)-(534, 60), RGB(128, 128, 128)
        End If

'If CurrentChannel = 0 Then
'        X = X + 3
'        Temp = 3
'End If
'If ChkChannel(Temp).Value = 1 Then
'    Picture1.Line (X, Y1(Temp))-(X + 3, Y2), MyColor(Temp)
'    'Y1(Temp) = Y2
'End If
End Sub

Private Sub Timer2_Timer()
PlotGraph
End Sub
