VERSION 5.00
Begin VB.Form vw 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��Ȩ��Ϣ"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7185
   Icon            =   "vw.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7185
   StartUpPosition =   1  '����������
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "δ��������ɣ��������أ�                                   ������ѧ����֪ʶ��Ȩ��"
      BeginProperty Font 
         Name            =   "������κ"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   5400
      Width           =   6855
   End
   Begin VB.Image Image3 
      Height          =   2655
      Left            =   4080
      Picture         =   "vw.frx":068A
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   2175
      Left            =   3960
      Picture         =   "vw.frx":33DC8
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   2685
      Left            =   120
      Picture         =   "vw.frx":37409
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   3705
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ԣ�����֣�  ����CJSoft��          ETB"
      BeginProperty Font 
         Name            =   "��Բ"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   1455
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������ҳ��http://home.chuyihao.com"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   150
      TabIndex        =   1
      Top             =   600
      Width           =   3570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��д by �����е���ʮ����ѧ CJSoft���ݺ�"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3510
   End
End
Attribute VB_Name = "vw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label1.Caption = "��д by �����е���ʮ����ѧ CJSoft���ݺ�"
End Sub

Public Sub Label2_Click()
Shell "explorer.exe http://home.chuyihao.com", vbNormalFocus
End Sub
