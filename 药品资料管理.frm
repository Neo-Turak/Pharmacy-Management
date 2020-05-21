VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form 药品信息修改 
   Caption         =   "药品资料管理"
   ClientHeight    =   9255
   ClientLeft      =   855
   ClientTop       =   1560
   ClientWidth     =   18255
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   18255
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   9960
      Top             =   8160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   794
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7335
      Left            =   5400
      TabIndex        =   24
      Top             =   480
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   12938
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
   Begin VB.Frame Frame1 
      Caption         =   "药品详细资料"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   4695
      Begin MSForms.CommandButton CommandButton1 
         Height          =   495
         Left            =   1560
         TabIndex        =   23
         Top             =   6000
         Width           =   1935
         Caption         =   "修改"
         Size            =   "3413;873"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.TextBox TextBox11 
         Height          =   375
         Left            =   1440
         TabIndex        =   22
         Top             =   5400
         Width           =   2535
         VariousPropertyBits=   746604571
         Size            =   "4471;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextBox10 
         Height          =   375
         Left            =   1440
         TabIndex        =   21
         Top             =   4920
         Width           =   2535
         VariousPropertyBits=   746604571
         Size            =   "4471;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextBox9 
         Height          =   375
         Left            =   1440
         TabIndex        =   20
         Top             =   4440
         Width           =   2535
         VariousPropertyBits=   746604571
         Size            =   "4471;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextBox8 
         Height          =   375
         Left            =   1440
         TabIndex        =   19
         Top             =   3960
         Width           =   2535
         VariousPropertyBits=   746604571
         Size            =   "4471;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextBox7 
         Height          =   375
         Left            =   1440
         TabIndex        =   18
         Top             =   3480
         Width           =   2535
         VariousPropertyBits=   746604571
         Size            =   "4471;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextBox6 
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   3000
         Width           =   2535
         VariousPropertyBits=   746604571
         Size            =   "4471;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextBox5 
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   2640
         Width           =   2535
         VariousPropertyBits=   746604571
         Size            =   "4471;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextBox4 
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   2160
         Width           =   2535
         VariousPropertyBits=   746604571
         Size            =   "4471;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextBox3 
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         Top             =   1560
         Width           =   2535
         VariousPropertyBits=   746604571
         Size            =   "4471;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextBox2 
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   1080
         Width           =   2535
         VariousPropertyBits=   746604571
         Size            =   "4471;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextBox1 
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   600
         Width           =   2535
         VariousPropertyBits=   746604571
         Size            =   "4471;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   375
         Index           =   10
         Left            =   360
         TabIndex        =   11
         Top             =   5400
         Width           =   1455
         Caption         =   "产地"
         Size            =   "2566;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   375
         Index           =   9
         Left            =   360
         TabIndex        =   10
         Top             =   4920
         Width           =   1455
         Caption         =   "生产商"
         Size            =   "2566;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   375
         Index           =   8
         Left            =   360
         TabIndex        =   9
         Top             =   4440
         Width           =   855
         Caption         =   "批号"
         Size            =   "1508;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   375
         Index           =   7
         Left            =   360
         TabIndex        =   8
         Top             =   3000
         Width           =   1455
         Caption         =   "药品类型"
         Size            =   "2566;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   375
         Index           =   6
         Left            =   360
         TabIndex        =   7
         Top             =   3480
         Width           =   1455
         Caption         =   "剂型"
         Size            =   "2566;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   6
         Top             =   3960
         Width           =   1455
         Caption         =   "规格"
         Size            =   "2566;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   2520
         Width           =   975
         Caption         =   "俗名简码"
         Size            =   "1720;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Top             =   2040
         Width           =   1455
         Caption         =   "药品俗名"
         Size            =   "2566;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   1560
         Width           =   1455
         Caption         =   "药品简码"
         Size            =   "2566;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
         Caption         =   "药品名称"
         Size            =   "2566;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1455
         Caption         =   "药品编码"
         Size            =   "2566;661"
         FontName        =   "宋体"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   495
      Left            =   3960
      TabIndex        =   27
      Top             =   480
      Width           =   1095
      Caption         =   "查找"
      Size            =   "1931;873"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TextBox12 
      Height          =   495
      Left            =   1920
      TabIndex        =   26
      Top             =   480
      Width           =   1935
      VariousPropertyBits=   746604571
      Size            =   "3413;873"
      FontName        =   "宋体"
      FontHeight      =   180
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   495
      Left            =   360
      TabIndex        =   25
      Top             =   480
      Width           =   1455
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2566;873"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "宋体"
      FontHeight      =   180
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "药品信息修改"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
