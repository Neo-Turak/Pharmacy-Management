VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form 药典设置 
   Caption         =   "药典设置"
   ClientHeight    =   8415
   ClientLeft      =   3450
   ClientTop       =   990
   ClientWidth     =   13035
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   13035
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Caption         =   "基本信息"
      Height          =   8175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin VB.Timer Timer1 
         Left            =   3480
         Top             =   5400
      End
      Begin VB.TextBox Text7 
         DataField       =   "药品类型"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1080
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   400
         Width           =   2775
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "自动整理"
         Height          =   615
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   6480
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "删除"
         Height          =   615
         Left            =   240
         TabIndex        =   18
         Top             =   6480
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "保存"
         Height          =   615
         Left            =   2280
         TabIndex        =   17
         Top             =   5760
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "添加"
         Height          =   615
         Left            =   240
         TabIndex        =   16
         Top             =   5760
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         DataField       =   "采购价"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """￥""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1080
         TabIndex        =   15
         Text            =   "Text6"
         Top             =   5040
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         DataField       =   "剂型"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   4440
         Width           =   2775
      End
      Begin VB.TextBox Text4 
         DataField       =   "规格"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1080
         TabIndex        =   13
         Text            =   "Text4"
         Top             =   3840
         Width           =   2775
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         DataField       =   "总代理商"
         DataSource      =   "Adodc1"
         Height          =   855
         Left            =   60
         TabIndex        =   12
         Top             =   2880
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"药典设置.frx":0000
      End
      Begin VB.TextBox Text3 
         DataField       =   "速记码"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         DataField       =   "通用名"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   960
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         DataField       =   "流水号"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   960
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "药品类别"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Top             =   450
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "采购价"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "剂    型"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "药库规格"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "生产企业"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   2590
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "速记码"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "通用名"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "流水号"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "药典设置.frx":008F
      Height          =   7575
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   13361
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "药品类型"
         Caption         =   "药品类型"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "通用名"
         Caption         =   "通用名"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "速记码"
         Caption         =   "速记码"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "总代理商"
         Caption         =   "总代理商"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "规格"
         Caption         =   "规格"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "剂型"
         Caption         =   "剂型"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "采购价"
         Caption         =   "采购价"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """￥""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2190.047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3105.071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2069.858
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2115.213
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2069.858
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7080
      Top             =   7920
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   "药典"
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
End
Attribute VB_Name = "药典设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mHZtoSM As cHztoSM

Private Sub Command4_Click()
If MsgBox("自动整理操作进行速记码的转换并同时写入保存。进行期间不能进行其他操作，请慎重操作！继续吗？", 64 + vbYesNo, "警告") = vbYes Then
Timer1.Interval = 100
End If
End Sub

Private Sub Form_Load()
DataGrid1.Width = Screen.Width - 2000
DataGrid1.Height = Screen.Height - 2000
Set mHZtoSM = New cHztoSM
    
    mHZtoSM.LoadLibFile App.Path & "\GB2312SM.Lib"
    If mHZtoSM.LoadLibSuccess = False Then Unload Me
    
    End Sub
       
Private Sub Form_Unload(Cancel As Integer)
Set mHZtoSM = Nothing
End Sub

Private Sub Text3_GotFocus()
Text3.Text = mHZtoSM.HZtoSMEx(Text2.Text)
End Sub

Private Sub Text3_LostFocus()
Text3.Text = mHZtoSM.HZtoSMEx(Text2.Text)
End Sub

Private Sub Timer1_Timer()
Adodc1.Recordset.MoveNext
Text2.SetFocus
Text3.SetFocus
RichTextBox1.SetFocus

If Adodc1.Recordset.EOF = True Then
Timer1.Interval = 0
MsgBox "结束了！"
End If

End Sub
