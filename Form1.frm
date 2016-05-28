VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmetb 
   AutoRedraw      =   -1  'True
   Caption         =   "CJSoft Easy Trace Board v.0.1 BETA      37自产 CJSoft 褚姜软件 37中学部  褚逸豪作品"
   ClientHeight    =   8055
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   12720
   DrawMode        =   1  'Blackness
   DrawWidth       =   10
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   12720
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog dopen 
      Left            =   8520
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "打开一个ETB涂鸦文档"
      Filter          =   "位图（*.bmp)|*.bmp"
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   9000
      Top             =   7440
   End
   Begin VB.Frame frmset 
      Caption         =   "白板笔功能"
      Height          =   735
      Left            =   3960
      TabIndex        =   6
      Top             =   7320
      Width           =   2025
      Begin VB.CommandButton cmdc 
         Caption         =   "颜色"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdcc 
         Caption         =   "取消"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   1920
         Width           =   1180
      End
      Begin VB.CheckBox chkline 
         Caption         =   "自动直线"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   405
         Left            =   1080
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog dfont 
      Left            =   8040
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar padhsll 
      Height          =   255
      LargeChange     =   100
      Left            =   0
      Max             =   1000
      SmallChange     =   30
      TabIndex        =   5
      Top             =   7080
      Visible         =   0   'False
      Width           =   12735
   End
   Begin VB.VScrollBar padsll 
      Height          =   7095
      LargeChange     =   100
      Left            =   12480
      Max             =   1000
      SmallChange     =   30
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComDlg.CommonDialog dsave 
      Left            =   7560
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.bmp"
      DialogTitle     =   "保存ETB涂鸦文档"
      FileName        =   "ETB文档.bmp"
      Filter          =   "位图（*.bmp）|*.bmp"
   End
   Begin VB.OptionButton dmode 
      Caption         =   "白板笔"
      Height          =   735
      Index           =   0
      Left            =   6000
      Picture         =   "Form1.frx":068A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7320
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   735
      LargeChange     =   10
      Left            =   0
      Max             =   100
      Min             =   1
      SmallChange     =   10
      TabIndex        =   2
      Top             =   7320
      Value           =   10
      Width           =   3975
   End
   Begin VB.OptionButton dmode 
      Caption         =   "橡皮擦"
      Height          =   735
      Index           =   1
      Left            =   6720
      Picture         =   "Form1.frx":0D7D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
      Width           =   735
   End
   Begin VB.PictureBox p 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   10
      Height          =   7095
      Left            =   0
      ScaleHeight     =   7035
      ScaleWidth      =   12435
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   3000
         Top             =   6720
      End
   End
   Begin VB.Menu mfile 
      Caption         =   "文件"
      Begin VB.Menu dk 
         Caption         =   "打开"
         Shortcut        =   ^O
      End
      Begin VB.Menu msave 
         Caption         =   "保存"
         Shortcut        =   ^S
      End
      Begin VB.Menu masave 
         Caption         =   "另存为"
         Shortcut        =   ^A
      End
      Begin VB.Menu l 
         Caption         =   "童锁:开"
      End
   End
   Begin VB.Menu mabout 
      Caption         =   "作者主页"
   End
End
Attribute VB_Name = "frmetb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ppp As Boolean
Dim asa As Boolean
Dim BRUSH As Boolean
Dim bbb As Boolean
Dim isopen As Boolean
Dim sa As weizhi
Dim sas As Boolean
Dim savepath As String
Dim times As Integer

Private Sub cmdc_Click()

dfont.ShowColor
Shape1.BackColor = dfont.Color
p.ForeColor = dfont.Color
End Sub

Private Sub cmdcc_Click()
frmset.Visible = False
End Sub



Private Sub dk_Click()
dopen.ShowOpen

If Trim(dopen.FileName) <> "" Then p.Picture = LoadPicture(dopen.FileName)
dopen.FileName = ""
End Sub

Private Sub dmode_Click(Index As Integer)
If dmode(0).Value = True Then BRUSH = False
If dmode(1).Value = True Then BRUSH = True
End Sub

Private Sub dmode_DblClick(Index As Integer)
If Index = 1 Then
 p.Cls
 dmode(0).Value = True
Else
frmset.Top = Me.ScaleHeight - frmset.Height
frmset.Left = dmode(0).Left - frmset.Width
frmset.Visible = True
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then Timer1.Enabled = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then Timer1.Enabled = False
End Sub

Private Sub Form_Load()
Dim ts
Open "C:\windows\etbsetting.cjs" For Append As #1
Print #1, "1"
Close #1
Open "C:\windows\etbsetting.cjs" For Input As #1
Line Input #1, tsl
If tsl = 1 Then
l.Caption = "童锁:开"
Else
l.Caption = "童锁:关"
End If
Close #1
times = 0
savepath = ""
asa = False
BRUSH = False
bbb = True
Me.WindowState = 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If l.Caption = "童锁:开" Then
If Timer2.Enabled = False Then
Timer2.Enabled = True
times = times + 1
Cancel = 1
Else
times = times + 1
If times < 3 Then Cancel = 1
End If
End If
End Sub

Private Sub Form_Resize()
p.Width = Me.ScaleWidth
p.Height = Me.ScaleHeight - 735

HScroll1.Top = Me.ScaleHeight - 735
dmode(0).Top = Me.ScaleHeight - 735
dmode(1).Top = Me.ScaleHeight - 735
dmode(1).Left = Me.ScaleWidth - 735
dmode(0).Left = Me.ScaleWidth - 735 - 735
HScroll1.Width = Me.ScaleWidth - 735 - 735 - 2025
frmset.Left = Me.ScaleWidth - 735 - 735 - 2025
frmset.Top = dmode(1).Top
End Sub

Private Sub HScroll1_Change()
p.DrawWidth = HScroll1.Value
End Sub

Private Sub mnew_Click()
Shell App.Path & "\" & App.EXEName, vbNormalFocus
End Sub

Private Sub l_Click()
Open "C:\windows\etbsetting.cjs" For Output As #1
If l.Caption = "童锁:开" Then
l.Caption = "童锁:关"
Print #1, "0"
Else
l.Caption = "童锁:开"
Print #1, "1"
End If
Close #1
End Sub

Private Sub mabout_Click()
vw.Label2_Click
End Sub

Private Sub masave_Click()
dsave.ShowSave
If Trim(dsave.FileName) <> "" Then SavePicture p.Image, dsave.FileName
savepath = dsave.FileName
dsave.FileName = ""
End Sub

Private Sub msave_Click()
If savepath = "" Then
dsave.ShowSave
If Trim(dsave.FileName) <> "" Then SavePicture p.Image, dsave.FileName
savepath = dsave.FileName
dsave.FileName = ""
Else
SavePicture p.Image, savepath
End If
End Sub

Private Sub p_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim w As Integer
w = 1
If chkline.Value = 1 Then
If sas = False Then
sa.X = X
sa.Y = Y
sas = True
GoTo ll
ElseIf sas = True Then
p.Line (sa.X, sa.Y)-(X, Y)
sas = False
End If
End If
ll:
p.PSet (X, Y)
If BRUSH = True Then p.PSet (X, Y), p.BackColor
ppp = True
End Sub

Private Sub p_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sas = False
If ppp = False Then Exit Sub
If BRUSH = True Then GoTo bru
Static a As Long
Static b As Long
If asa = False Then
asa = True
p.Line (X, Y)-(X, Y)
a = X: b = Y
Exit Sub
End If
p.Line (a, b)-(X, Y)
a = X: b = Y
bru:
Static c As Long
Static d As Long

If bbb = True Then
c = X
d = Y
End If
If Button = 1 And BRUSH = True Then
p.Line (c, d)-(X, Y), p.BackColor
bbb = False
c = X
d = Y
End If
End Sub

Private Sub p_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ppp = False
asa = False
bbb = True
End Sub

Private Sub padhsll_Change()
p.Left = 0 - padhsll.Value * 10
p.Width = padhsll.Value * 10 + Me.ScaleWidth - 255
End Sub

Private Sub padhsll_Scroll()
p.Left = 0 - padhsll.Value * 10
p.Width = padhsll.Value * 10 + Me.ScaleWidth - 255
End Sub

Private Sub padsll_Change()
p.Top = 0 - padsll.Value * 10
p.Height = padsll.Value * 10 + Me.ScaleHeight - 735 - 255
End Sub

Private Sub padsll_Scroll()
p.Top = 0 - padsll.Value * 10
p.Height = padsll.Value * 10 + Me.ScaleHeight - 735 - 255
End Sub

Private Sub Timer1_Timer()
vw.Show
Timer1.Enabled = False
End Sub

Private Sub viewc_Click()
If viewc.Caption = "进入到无边框模式" Then
Me.BorderStyle = 0
viewc.Caption = "退出到有边框模式"
Else
Me.BorderStyle = 2
viewc.Caption = "进入到无边框模式"
End If
End Sub


