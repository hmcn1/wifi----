VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ȵ㴴��"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   3060
   BeginProperty Font 
      Name            =   "����"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleMode       =   0  'User
   ScaleWidth      =   3060
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox CODE 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox SSID 
      BeginProperty Font 
         Name            =   "��Բ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ�ϴ���"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "��Բ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "����: "
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
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "SSID:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
A = SSID.Text
B = CODE.Text
C = Len(B)
D = Len(A)

If D = 0 Then
    temp = MsgBox("�ȵ����Ʋ���Ϊ�գ�", , "����")
Else
    If C < 8 Then
        temp = MsgBox("���볤�Ȳ���С��8λ��", , "����")
    Else
        Shell "cmd.exe /k netsh wlan set hostednetwork mode=allow ssid=" & A & " key=" & B & "", vbHide
        Shell "cmd.exe /k netsh wlan start hostednetwork", vbHide
        temp = MsgBox("��������Wifi�ȵ�:" & vbCrLf & "   ����Ϊ:" & A & vbCrLf & "   ����Ϊ:" & B & "", , "�����ɹ���")
        Unload Me
    End If
End If
End Sub

