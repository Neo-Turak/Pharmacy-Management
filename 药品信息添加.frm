VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form ҩƷ��Ϣ��� 
   Caption         =   "ҩƷ�ֵ����"
   ClientHeight    =   10275
   ClientLeft      =   3255
   ClientTop       =   2670
   ClientWidth     =   14265
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10275
   ScaleWidth      =   14265
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "ҩƷ��Ϣ���.frx":0000
      Height          =   7095
      Left            =   120
      TabIndex        =   23
      Top             =   2760
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   12515
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "ҩƷID"
         Caption         =   "ҩƷID"
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
         DataField       =   "ҩƷ����"
         Caption         =   "ҩƷ����"
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
         DataField       =   "����"
         Caption         =   "����"
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
         DataField       =   "���"
         Caption         =   "���"
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
         DataField       =   "ҩƷ����"
         Caption         =   "ҩƷ����"
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
         DataField       =   "����"
         Caption         =   "����"
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
         DataField       =   "����"
         Caption         =   "����"
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
         DataField       =   "������"
         Caption         =   "������"
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
      BeginProperty Column08 
         DataField       =   "��ַ"
         Caption         =   "��ַ"
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
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2775.118
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   12240
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ҩƷ��Ϣ��"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   735
      Left            =   11040
      TabIndex        =   22
      Top             =   2040
      Width           =   975
      Size            =   "1720;1296"
      Picture         =   "ҩƷ��Ϣ���.frx":0015
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TextBox9 
      Height          =   375
      Left            =   5400
      TabIndex        =   21
      Top             =   2040
      Width           =   2535
      VariousPropertyBits=   746604571
      Size            =   "4471;661"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox8 
      Height          =   375
      Left            =   1320
      TabIndex        =   20
      Top             =   2040
      Width           =   2535
      VariousPropertyBits=   746604571
      Size            =   "4471;661"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox7 
      Height          =   375
      Left            =   5400
      TabIndex        =   19
      Top             =   1440
      Width           =   2535
      VariousPropertyBits=   746604571
      Size            =   "4471;661"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox6 
      Height          =   375
      Left            =   9720
      TabIndex        =   18
      Top             =   1440
      Width           =   2535
      VariousPropertyBits=   746604571
      Size            =   "4471;661"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox ComboBox2 
      Height          =   375
      Left            =   1320
      TabIndex        =   17
      Top             =   1440
      Width           =   2535
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4471;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   375
      Left            =   9720
      TabIndex        =   16
      Top             =   840
      Width           =   2535
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4471;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox5 
      Height          =   375
      Left            =   5400
      TabIndex        =   15
      Top             =   840
      Width           =   2535
      VariousPropertyBits=   746604571
      Size            =   "4471;661"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox4 
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Top             =   840
      Width           =   2535
      VariousPropertyBits=   746604571
      Size            =   "4471;661"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox3 
      Height          =   375
      Left            =   9720
      TabIndex        =   13
      Top             =   240
      Width           =   2535
      VariousPropertyBits=   746604571
      Size            =   "4471;661"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox2 
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   240
      Width           =   2535
      VariousPropertyBits=   746604571
      Size            =   "4471;661"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   240
      Width           =   2535
      VariousPropertyBits=   746604575
      Size            =   "4471;661"
      Value           =   "<�Զ�����>"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Index           =   10
      Left            =   4320
      TabIndex        =   10
      Top             =   2040
      Width           =   735
      Caption         =   "��ַ"
      Size            =   "1296;661"
      FontName        =   "����"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Index           =   9
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   1575
      Caption         =   "������"
      Size            =   "2778;661"
      FontName        =   "����"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Index           =   8
      Left            =   4320
      TabIndex        =   8
      Top             =   1440
      Width           =   855
      Caption         =   "����"
      Size            =   "1508;661"
      FontName        =   "����"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Index           =   7
      Left            =   8880
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
      Caption         =   "���"
      Size            =   "2143;661"
      FontName        =   "����"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
      Caption         =   "����"
      Size            =   "2778;661"
      FontName        =   "����"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Index           =   5
      Left            =   8640
      TabIndex        =   5
      Top             =   840
      Width           =   1575
      Caption         =   "ҩƷ����"
      Size            =   "2778;661"
      FontName        =   "����"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Index           =   4
      Left            =   4200
      TabIndex        =   4
      Top             =   840
      Width           =   1095
      Caption         =   "��������"
      Size            =   "1931;661"
      FontName        =   "����"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1575
      Caption         =   "ҩƷ����"
      Size            =   "2778;661"
      FontName        =   "����"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Index           =   2
      Left            =   8640
      TabIndex        =   2
      Top             =   240
      Width           =   1575
      Caption         =   "���Ƽ���"
      Size            =   "2778;661"
      FontName        =   "����"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   1575
      Caption         =   "ҩƷ����"
      Size            =   "2778;661"
      FontName        =   "����"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
      Caption         =   "ҩƷ����"
      Size            =   "2778;661"
      FontName        =   "����"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
End
Attribute VB_Name = "ҩƷ��Ϣ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
