VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wifi�ȵ㹤��v1.4"
   ClientHeight    =   4530
   ClientLeft      =   8700
   ClientTop       =   3840
   ClientWidth     =   5490
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   5490
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "��Բ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ȡ���ػ�"
      BeginProperty Font 
         Name            =   "��Բ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "��Բ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1920
      TabIndex        =   6
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   2400
      TabIndex        =   7
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "IP��ַ���·���"
      BeginProperty Font 
         Name            =   "��Բ"
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
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����Wifi�ȵ�"
      BeginProperty Font 
         Name            =   "��Բ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ʹ�ý̳�"
      BeginProperty Font 
         Name            =   "��Բ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�ر��ȵ�"
      BeginProperty Font 
         Name            =   "��Բ"
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
      Caption         =   "���ȵ�"
      BeginProperty Font 
         Name            =   "��Բ"
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
   Begin VB.Label Label7 
      BackColor       =   &H80000004&
      Caption         =   "--------Copyright By Wesley.H--------"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   3120
      TabIndex        =   18
      Top             =   4320
      Width           =   3375
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000010&
      X1              =   4680
      X2              =   5400
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000010&
      X1              =   3360
      X2              =   3720
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label6 
      Caption         =   "�ȵ����"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   720
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
      Y1              =   840
      Y2              =   4080
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
      Index           =   2
      X1              =   0
      X2              =   240
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   240
      Y1              =   0
      Y2              =   0
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
      Caption         =   "���Ӻ�ػ�"
      BeginProperty Font 
         Name            =   "��Բ"
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
      TabIndex        =   16
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
      Caption         =   "��ʱ��ػ�"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   15
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "����ʱ�ػ�"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "��"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "�ػ�"
      BeginProperty Font 
         Name            =   "��Բ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "���Թ���Ա������г��򣡱������������Win8"
      BeginProperty Font 
         Name            =   "����"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Unload Form2
If App.PrevInstance Then
MsgBox "���Ѿ�����������ˣ�"
End
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form2
End
End Sub

Private Sub Command1_Click()
Shell "cmd.exe /k netsh wlan start hostednetwork", vbHide
Shell "cmd.exe /k ipconfig /renew", vbHide
End Sub

Private Sub Command2_Click()
Shell "cmd.exe /k netsh wlan stop hostednetwork", vbHide
End Sub

Private Sub Command3_Click()
Form2.Show
End Sub

Private Sub Command4_Click()
If Len(Dir(App.Path & "\ʹ�ý̳�.txt")) = 0 Then
A = MsgBox("�̳��ļ��Ѳ����ڣ���������Ҫ����ϵ���ߣ�", , "����")
Else
Shell "cmd.exe /c start " & App.Path & "\ʹ�ý̳�.txt", vbHide
End If
End Sub

Private Sub Command5_Click()
Form3.Show
End Sub

Private Sub Command6_Click()
Shell "cmd.exe /k ipconfig /renew", vbHide
End Sub

Private Sub Command7_Click()
T = Text1.Text
p = IsNumeric(T)
If Len(T) = 0 Then
A = MsgBox("������һ��ʱ�䣡", , "����")
Text1.Text = ""
Else
If p = 0 Then
A = MsgBox("���������֣�", , "����")
Text1.Text = ""
Else
If T = 0 Then
A = MsgBox("ʱ��Ϊ0������̹ػ�������������ʱ��ֵ��", , "����")
Text1.Text = ""
Else
If T <= 0 Then
A = MsgBox("������һ��������", , "����")
Text1.Text = ""
Else

T1 = T * 60
F = Fix(T)
Shell "cmd.exe /k shutdown -s -hybrid -t " & T1 & "", vbHide
A = MsgBox("        ϵͳ����" & F & "���Ӻ�ػ�" + vbCrLf + "���������趨�ػ�ʱ������ȡ����ʱ�ػ�!", , "���ö�ʱ�ػ��ɹ���")
Text1.Text = ""
Text1.SetFocus
End If
End If
End If
End If

End Sub

Private Sub Command8_Click()
Shell "cmd.exe /k shutdown -a", vbHide
Shell "cmd.exe /k SCHTASKS /Delete /TN ��ʱ�ػ� /f", vbHide
MsgBox ("ȡ���ػ��ɹ�")
End Sub

Private Sub Command9_Click()
Dim Time As Long
L = Text2.Text
R = Text3.Text
Z1 = IsNumeric(L)
Z2 = IsNumeric(R)

If Z1 = 0 Or Z2 = 0 Then
    A = MsgBox("������һ��ʱ�䣡", , "����")
Else
    If L < 0 Or L >= 24 Then
        A = MsgBox("�������������0��23֮��������", , "����")
        Text2.SetFocus
    Else
        If R < 0 Or R >= 60 Then
            A = MsgBox("�����ҿ�������0��59֮��������", , "����")
            Text3.SetFocus
        Else
            L = Fix(L)
            R = Fix(R)
            H = Hour(Now)
            M = Minute(Now)
            S = Second(Now)
            S = S + 1
            
            If L = H And R = M Then
                A = MsgBox("���뵱ǰʱ��������̹ػ�������������ʱ�䣡", , "����")
            Else
                If L > H Or (L = H And R > M) Then
                    Time = 3600 * (L - H) + 60 * (R - M)
                Else
                    If L < H Or (L = H And R < M) Then
                        Time = 3600 * (L - H + 24) + 60 * (R - M)
                    End If
                End If
                A = MsgBox("ϵͳ��������" & L & "��" & R & "��" & S & "��ػ�" + vbCrLf + "���������趨�ػ�ʱ������ȡ����ʱ�ػ�!", , "���ö�ʱ�ػ��ɹ���")
                Shell "cmd.exe /k shutdown -s -hybrid -t " & Time & "", vbHide
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

