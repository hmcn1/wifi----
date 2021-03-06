VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wifi热点工具V2.0正式版"
   ClientHeight    =   4530
   ClientLeft      =   8700
   ClientTop       =   3840
   ClientWidth     =   5490
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   5490
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   -840
      TabIndex        =   0
      Top             =   -860
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   3720
   End
   Begin VB.CommandButton Command5 
      Caption         =   "查看已连接设备"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "取消关机"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1920
      TabIndex        =   7
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   2400
      TabIndex        =   8
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "IP地址重新分配"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "创建Wifi热点"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "使用教程"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "关闭热点"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开热点"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "请以管理员身份运行程序！本程序仅适用于Win8"
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000004&
      Caption         =   "--------Copyright By Wesley.H--------"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   3120
      TabIndex        =   19
      Top             =   4320
      Width           =   3375
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000010&
      X1              =   4680
      X2              =   5400
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000010&
      X1              =   3360
      X2              =   3720
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label6 
      Caption         =   "热点管理"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   600
      Width           =   735
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000010&
      X1              =   3360
      X2              =   3000
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   3360
      X2              =   3360
      Y1              =   720
      Y2              =   4200
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   1440
      X2              =   1920
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   360
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "分钟后关机"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   17
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   1680
      X2              =   1680
      Y1              =   3600
      Y2              =   1800
   End
   Begin VB.Label Label2 
      Caption         =   "按时间关机"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   16
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "倒计时关机"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "关机"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   2280
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean
Private Sub Form_Load()
Unload Form2
If App.PrevInstance Then
    MsgBox "你已经打开这个程序了！"
End
End If
Open "C:\Windows\temp3.txt" For Output As #1
Close #1
flag = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form2
Kill "C:\Windows\temp3.txt"
End
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
Shell "cmd.exe /c arp -a -N 192.168.137.1 >C:\Windows\temp3.txt", vbHide
Text4.SetFocus
Dim times As Single
times = Timer + 0.1
Do
    DoEvents
Loop While times > Timer
If FileLen("C:\Windows\temp3.txt") <> 0 Then
    MsgBox "您的热点已打开，无需重复打开", , "错误！"
    Text4.SetFocus
Else
    Shell "cmd.exe /c netsh wlan start hostednetwork", vbHide
End If
Shell "cmd.exe /c ipconfig /renew", vbHide
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Command2.Enabled = False
Text4.SetFocus
Shell "cmd.exe /c netsh wlan stop hostednetwork", vbHide
Command2.Enabled = True
End Sub

Private Sub Command3_Click()
Form2.Show
End Sub

Private Sub Command4_Click()
If Len(Dir(App.Path & "\使用教程.txt")) = 0 Then
    A = MsgBox("教程文件已不存在！！如有需要请联系作者！", , "错误！")
Else
    Shell "cmd.exe /c start " & App.Path & "\使用教程.txt", vbHide
End If
End Sub

Private Sub Command5_Click()
Shell "cmd.exe /c arp -a -N 192.168.137.1 >C:\Windows\temp.txt", vbHide
Shell "cmd.exe /c netsh wlan show hostednetwork >C:\Windows\temp2.txt", vbHide
Command5.Enabled = False
Text4.SetFocus
Head = "                  IP地址                        MAC地址"
Dim times As Single
times = Timer + 0.2
Do
    DoEvents
Loop While times > Timer
If FileLen("C:\Windows\temp.txt") = 0 Then
    MsgBox "热点未打开或IP地址未设置为192.168.137.1！", , "错误！"
Else
    Dim Str(1 To 50), Sw(1 To 50) As String
    Dim Si1() As String
    Dim Si2() As String
    Dim Si3() As String
    Dim Si, Sx As String
    Str(1) = "空"
    Open "C:\Windows\temp2.txt" For Input As #1
    Open "C:\Windows\temp.txt" For Input As #2
    i = 1
    Do Until EOF(1)
        Line Input #1, Sx
        If Right(Sx, 7) = "已经过身份验证" Then
            Si1 = Split(Sx, "        ", 3)
            Si3 = Split(Si1(1), ":", 6)
            Sw(i) = Si3(0) & "-" & Si3(1) & "-" & Si3(2) & "-" & Si3(3) & "-" & Si3(4) & "-" & Si3(5)
            i = i + 1
        End If
    Loop
    i = 1
    Do Until EOF(2)
        Line Input #2, Si
        If Right(Si, 32) = "" & Sw(i) & "     动态        " Then
            Si2 = Split(Si, "     ", 3)
            Si2(1) = UCase(Si2(1))
            Si2(0) = "设备" & i & "：" & Si2(0)
            Str(i) = Si2(0) & "       " & Si2(1)
            i = i + 1
            Seek #2, 1
        End If
    Loop
    Close #1
    Close #2
    If Str(1) = "空" Then
        MsgBox "暂无已连接设备！", , "已连接设备"
    Else
        For m = 1 To 10
            St = St & Str(m) & Chr(10) & Chr(13)
        Next m
        MsgBox "" & Head & "" + vbCrLf + vbCrLf + "" & St & "", , "已连接设备共" & i - 1 & "个"""
    End If
End If
Kill "C:\Windows\temp.txt"
Kill "C:\Windows\temp2.txt"
Command5.Enabled = True
End Sub

Private Sub Command6_Click()
Shell "cmd.exe /c ipconfig /renew", vbHide
End Sub

Private Sub Command7_Click()
T = Text1.Text
p = IsNumeric(T)
If Len(T) = 0 Then
    A = MsgBox("请输入一个时间！", , "错误！")
Else
    If p = 0 Then
        A = MsgBox("请输入数字！", , "错误！")
    Else
        If T = 0 Then
            A = MsgBox("时间为0则会立刻关机，请输入其他时间值！", , "错误！")
        Else
            If T <= 0 Then
                A = MsgBox("请输入一个正数！", , "错误！")
            Else
                T1 = T * 60
                F = Fix(T)
                Shell "cmd.exe /c shutdown -s -hybrid -t " & T1 & "", vbHide
                A = MsgBox("        系统将在" & F & "分钟后关机" + vbCrLf + "若需重新设定关机时间请先取消定时关机!", , "设置定时关机成功！")
            End If
        End If
    End If
End If
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub Command8_Click()
Shell "cmd.exe /c shutdown -a", vbHide
MsgBox ("取消关机成功")
End Sub

Private Sub Command9_Click()
Dim Time As Long
Dim WZ As String
L = Text2.Text
R = Text3.Text
Z1 = IsNumeric(L)
Z2 = IsNumeric(R)

If Z1 = 0 Or Z2 = 0 Then
    A = MsgBox("请输入一个时间！", , "错误！")
Else
    If L < 0 Or L >= 24 Then
        A = MsgBox("请在左框内输入0到23之间整数！", , "错误！")
        Text2.SetFocus
    Else
        If R < 0 Or R >= 60 Then
            A = MsgBox("请在右框内输入0到59之间整数！", , "错误！")
            Text2.SetFocus
        Else
            L = Fix(L)
            R = Fix(R)
            H = Hour(Now)
            m = Minute(Now)
            S = Second(Now)
            S = S + 1
            
            If L = H And R = m Then
                A = MsgBox("输入当前时间则会立刻关机，请重新输入时间！", , "错误！")
                Text2.SetFocus
            Else
                If L > H Or (L = H And R > m) Then
                    WZ = "今"
                Else
                    If L < H Or (L = H And R < m) Then
                        H = H - 24
                        WZ = "明"
                    End If
                End If
                Time = 3600 * (L - H) + 60 * (R - m)
                A = MsgBox("系统将在" & WZ & "天" & L & "点" & R & "分" & S & "秒关机" + vbCrLf + "若需重新设定关机时间请先取消定时关机!", , "设置定时关机成功！")
                Shell "cmd.exe /c shutdown -s -hybrid -t " & Time & "", vbHide
            End If
        End If
    End If
    Text2.Text = ""
    Text3.Text = ""
End If
End Sub

Sub Text1_keyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    Command7_Click
End If
End Sub
Sub Text2_keyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    Text3.SetFocus
End If
End Sub
Sub Text3_keyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    Command9_Click
End If
End Sub

Private Sub Timer1_Timer()
If flag = True Then
    Label8.Left = Label8.Left + 50
    If Label8.Left > Me.Width - Label8.Width Then flag = False
    Else
        Label8.Left = Label8.Left - 50
    If Label8.Left < 0 Then flag = True
End If
End Sub
