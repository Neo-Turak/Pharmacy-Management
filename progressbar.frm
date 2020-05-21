VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form progressbar 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   0
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   380
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "操作进行中"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "progressbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
ProgressBar1.Value = 0
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value = 10 Then
Label1.Caption = "操作进行中."
End If
If ProgressBar1.Value = 20 Then
Label1.Caption = "操作进行中.."
End If
If ProgressBar1.Value = 30 Then
Label1.Caption = "操作进行中..."
End If

If ProgressBar1.Value = 40 Then
Label1.Caption = "操作进行中."
End If
If ProgressBar1.Value = 50 Then
Label1.Caption = "操作进行中.."
End If
If ProgressBar1.Value = 60 Then
Label1.Caption = "操作进行中..."
End If

If ProgressBar1.Value = 70 Then
Label1.Caption = "操作进行中."
End If
If ProgressBar1.Value = 80 Then
Label1.Caption = "操作进行中.."
End If
If ProgressBar1.Value = 90 Then
Label1.Caption = "操作进行中..."
End If
If ProgressBar1.Value = 100 Then
Timer1.Enabled = False
Me.Hide
End If
End Sub
