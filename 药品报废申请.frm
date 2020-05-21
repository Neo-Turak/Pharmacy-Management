VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form 报废申请 
   Caption         =   "药品报废申请单"
   ClientHeight    =   7320
   ClientLeft      =   5010
   ClientTop       =   3300
   ClientWidth     =   6735
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   6735
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1508
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSForms.TextBox TextBox3 
      Height          =   495
      Left            =   1680
      TabIndex        =   12
      Top             =   3840
      Width           =   1335
      VariousPropertyBits=   746604571
      Size            =   "2355;873"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label5 
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   11
      Top             =   3960
      Width           =   975
      Caption         =   "报废数量"
      Size            =   "1720;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   615
      Left            =   2040
      TabIndex        =   9
      Top             =   6360
      Width           =   1575
      Caption         =   "申请"
      Size            =   "2778;1085"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   5280
      Width           =   2655
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4683;873"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label5 
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   7
      Top             =   5400
      Width           =   1335
      Caption         =   "处理措施"
      Size            =   "2355;450"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox2 
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   4560
      Width           =   3255
      VariousPropertyBits=   746604571
      Size            =   "5741;873"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
      Caption         =   "报废原因"
      Size            =   "2143;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
      Caption         =   "药品详细信息"
      Size            =   "2778;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   615
      Left            =   5040
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
      Caption         =   "查询"
      Size            =   "2143;1085"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1680
      Width           =   2895
      VariousPropertyBits=   746604571
      Size            =   "5106;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
      Caption         =   "药品名称"
      Size            =   "2355;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   3375
      Caption         =   "药品报废申请单"
      Size            =   "5953;873"
      FontName        =   "微软雅黑"
      FontHeight      =   435
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "报废申请"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
