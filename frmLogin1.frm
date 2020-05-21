VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLogin1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "登录"
   ClientHeight    =   3795
   ClientLeft      =   2790
   ClientTop       =   3105
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin1.frx":0000
   ScaleHeight     =   2242.211
   ScaleMode       =   0  'User
   ScaleWidth      =   3056.268
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   600
      Top             =   3840
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dbase\43.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dbase\43.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "用户表"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin MSForms.CommandButton CommandButton2 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   480
      TabIndex        =   5
      Top             =   3000
      Width           =   2295
      ForeColor       =   -2147483637
      VariousPropertyBits=   19
      Caption         =   "登录"
      Size            =   "4048;741"
      FontName        =   "宋体"
      FontHeight      =   285
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "药库工作站"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "password"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "用户名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   4800
      Width           =   495
   End
   Begin MSForms.TextBox TxtPassword 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
      VariousPropertyBits=   746604563
      BorderStyle     =   1
      Size            =   "3413;661"
      PasswordChar    =   42
      SpecialEffect   =   0
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextUserName 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   1935
      VariousPropertyBits=   746604563
      BorderStyle     =   1
      Size            =   "3413;873"
      SpecialEffect   =   0
      FontName        =   "宋体"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   134
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmLogin1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conn As ADODB.connection
Dim mrc As ADODB.Recordset
Option Explicit
Public LoginSucceeded As Boolean
Private Sub cmdCancel_Click()
End
End Sub

Private Sub CommandButton2_Click()
connect (TextUserName.Text)
Set Adodc1.Recordset = mrc
If Label4.Caption = TxtPassword.Text Then
        信息提示器.Show
        Me.Hide
    '设置全局变量为 false
    '不提示失败的登录
    Else
    LoginSucceeded = False
    MsgBox "用户名或者密码错误，请重试！", vbDefaultButton1, "登录错误！"
    End If
End Sub

Private Sub Form_Load()
Dim x As Integer, y As Integer
x = Screen.Width / Screen.TwipsPerPixelX
y = Screen.Height / Screen.TwipsPerPixelY
End Sub


Public Sub connect(username As String)
Set Conn = New ADODB.connection
Set mrc = New ADODB.Recordset
Dim constring As String
constring = "Provider=Microsoft.jet.OLEDB.4.0;Data Source=" & App.Path & "\dbase\43.mdb;Persist Security Info=False"
Conn.Open constring
Conn.CursorLocation = adUseClient
mrc.Open "select password from 用户表 where username='" & username & "'", Conn, adOpenKeyset, adLockOptimistic
End Sub

Private Sub Label1_Click()
End
End Sub
