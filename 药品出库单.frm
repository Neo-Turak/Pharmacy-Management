VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form ҩƷ���� 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ҩƷ���ڵ�"
   ClientHeight    =   10170
   ClientLeft      =   2535
   ClientTop       =   1905
   ClientWidth     =   12075
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   10170
   ScaleWidth      =   12075
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "��ӡ��ǰ"
      Height          =   375
      Left            =   8520
      TabIndex        =   27
      Top             =   9360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��"
      Height          =   375
      Left            =   8040
      TabIndex        =   26
      Top             =   9360
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   3960
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "ҩƷ���ⵥ.frx":0000
      Height          =   4695
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "���ⵥ��"
         Caption         =   "���ⵥ��"
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
         DataField       =   "ȡҩ��λ"
         Caption         =   "ȡҩ��λ"
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
         DataField       =   "ҩƷ��"
         Caption         =   "ҩƷ��"
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      BeginProperty Column06 
         DataField       =   "�ۼ�"
         Caption         =   "�ۼ�"
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
         DataField       =   "�ܼ�"
         Caption         =   "�ܼ�"
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
         DataField       =   "����ʱ��"
         Caption         =   "����ʱ��"
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
      BeginProperty Column09 
         DataField       =   "��Ч��"
         Caption         =   "��Ч��"
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
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1665.071
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1920.189
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   3960
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      ItemData        =   "ҩƷ���ⵥ.frx":0015
      Left            =   4200
      List            =   "ҩƷ���ⵥ.frx":0025
      TabIndex        =   5
      Top             =   3960
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   120
      Top             =   9360
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=logs.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=logs.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "����"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   9600
      Top             =   4080
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=logs.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=logs.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "���"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "ҩƷ���ⵥ.frx":0043
      Height          =   2295
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4048
      _Version        =   393216
      BackColor       =   -2147483632
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
      ColumnCount     =   13
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
         DataField       =   "��ⵥ��"
         Caption         =   "��ⵥ��"
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
         DataField       =   "�ۼ�"
         Caption         =   "�ۼ�"
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
      BeginProperty Column09 
         DataField       =   "�ܼ�"
         Caption         =   "�ܼ�"
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
      BeginProperty Column10 
         DataField       =   "��������"
         Caption         =   "��������"
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
      BeginProperty Column11 
         DataField       =   "��Ч����"
         Caption         =   "��Ч����"
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
      BeginProperty Column12 
         DataField       =   "�������"
         Caption         =   "�������"
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
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1305.071
         EndProperty
      EndProperty
   End
   Begin VB.Label Label18 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """��""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   2
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   25
      Top             =   9360
      Width           =   1695
   End
   Begin VB.Label Label17 
      Caption         =   "��ֵ��"
      Height          =   375
      Left            =   5400
      TabIndex        =   24
      Top             =   9360
      Width           =   735
   End
   Begin VB.Label Label16 
      Height          =   375
      Left            =   3720
      TabIndex        =   23
      Top             =   9360
      Width           =   1095
   End
   Begin VB.Label Label15 
      Caption         =   "��������"
      Height          =   375
      Left            =   2640
      TabIndex        =   22
      Top             =   9360
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "���ⵥ��"
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FF8080&
      Caption         =   "13"
      DataField       =   "��Ч����"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   10440
      TabIndex        =   20
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FF8080&
      Caption         =   "12"
      DataField       =   "�ܼ�"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9240
      TabIndex        =   19
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FF8080&
      Caption         =   "11"
      DataField       =   "����"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   8640
      TabIndex        =   18
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FF8080&
      Caption         =   "10"
      DataField       =   "�ۼ�"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7920
      TabIndex        =   17
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF8080&
      Caption         =   "9"
      DataField       =   "����"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7200
      TabIndex        =   16
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      Caption         =   "8"
      DataField       =   "������"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4800
      TabIndex        =   15
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   " ����������"
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      Caption         =   "6"
      DataField       =   "���"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      Caption         =   "5"
      DataField       =   "����"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "4"
      DataField       =   "����"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   2175
   End
   Begin MSForms.Label Label3 
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   3960
      Width           =   975
      Caption         =   "���⵽"
      Size            =   "1720;661"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   495
      Left            =   8400
      TabIndex        =   7
      Top             =   3960
      Width           =   1095
      Caption         =   "����"
      Size            =   "1931;873"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   855
      Caption         =   "����"
      Size            =   "1508;873"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   2295
      VariousPropertyBits=   746604571
      Size            =   "4048;661"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   360
      Width           =   975
      Caption         =   "ҩƷ����"
      Size            =   "1720;661"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   240
      Width           =   1815
      Caption         =   "ҩƷ����"
      Size            =   "3201;873"
      FontName        =   "��Բ"
      FontHeight      =   435
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "ҩƷ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo err1:

Dim cnt As Integer
Dim Amont As Integer

Dim price As Currency
Dim Profit As Currency

cnt = Adodc1.Recordset.RecordCount

If cnt = 0 Then
Label16.Caption = 0
Label18.Caption = 0
Else
Adodc2.Recordset.MoveFirst
Do Until Adodc2.Recordset.EOF = True
price = price + Adodc2.Recordset.Fields("�ܼ�").Value
Amont = Amont + Adodc2.Recordset.Fields("����").Value
Adodc2.Recordset.MoveNext
Loop
Label18.Caption = price
Label16.Caption = Amont
End If
Exit Sub
err1:
MsgBox Err.Number & Err.Description
End Sub

Private Sub Command2_Click()
Call Command1_Click
Printer.PaperSize = 9
'��ʼ��ӡ����
'���жԻ�����������
Printer.Orientation = 1
Printer.ScaleWidth = 21
Printer.ScaleHeight = 27
Printer.FontBold = True
Printer.FontItalic = True
Printer.FontSize = 16
Printer.CurrentX = 10
Printer.CurrentY = 2
Printer.Print "���ⵥ"      'Title

Printer.FontBold = False
Printer.FontItalic = False
Printer.FontSize = 12
 Adodc2.Recordset.MoveFirst
 Dim rt As ADODB.Recordset
 Set rt = Adodc2.Recordset
 With rt
 n = 3
Do Until rt.EOF = True

Printer.CurrentX = 1
Printer.CurrentY = n

 Printer.Print .Fields(0).Value & vbTab & .Fields(1).Value & vbTab & .Fields(2).Value & vbTab & .Fields(3).Value & vbTab & .Fields(4).Value & vbTab & .Fields(5).Value & _
   vbTab & .Fields(6).Value & vbTab & .Fields(7).Value & Space(2) & .Fields(8).Value
 rt.MoveNext
 n = n + 1
 Loop
 Printer.Print vbTab & "----------------------------------------------------------------------"
 Printer.Print vbTab & "������" & vbTab & Label16.Caption & vbTab & "��ֵ����" & Label18.Caption
 Printer.EndDoc
 
End With
End Sub

Private Sub CommandButton1_Click()
On Error GoTo err1:
Dim SS As String
Dim conn As ADODB.connection
Dim mrc As ADODB.Recordset
  Set conn = New ADODB.connection
  Set mrc = New ADODB.Recordset
Dim ConnectString As String
ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\logs.mdb;Persist Security Info=False"
'*������
conn.Open ConnectString
'*�����α�λ��
conn.CursorLocation = adUseClient
mrc.Open "select * from ��� where ���� like '" & TextBox1.Text & "%'", conn, 3, 2
Set Adodc1.Recordset = mrc
Set DataGrid1.DataSource = mrc
Exit Sub
err1:
MsgBox "���ִ���" & vbCrLf & Err.Number & Err.Description, vbOKOnly, "ERROR"
End Sub


Private Sub CommandButton2_Click()
On Error GoTo err1:
Dim rs As ADODB.Recordset
Set rs = Adodc2.Recordset
If Label4.Caption <> "" Then
progressbar.Show
  rs.AddNew
  rs!���ⵥ�� = Text2.Text
  rs!ȡҩ��λ = Combo1.Text
  rs!���� = Text1.Text
  rs!���� = Label5.Caption
  rs!��� = Label6.Caption
  rs!�ۼ� = Label10.Caption
  rs!�ܼ� = Label10.Caption * Text1.Text
  rs!����ʱ�� = Date
  rs!��Ч�� = Label13.Caption
  rs!ҩƷ�� = Label4.Caption
  rs.update
  Dim qm As Integer
  
qm = Val(Label11.Caption) - Val(Text1.Text)
'this line for change the result value
Dim rt As ADODB.Recordset
Set rt = Adodc1.Recordset
rt.update
rt!���� = qm
rt.update
End If
Exit Sub
err1:
MsgBox "���ִ���" & vbCrLf & Err.Number & Err.Description, vbOKOnly, "ERROR"
End Sub

Private Sub Form_Load()
DataGrid1.Width = Screen.Width - 500
DataGrid2.Width = Screen.Width - 500
End Sub

Private Sub Text1_LostFocus()
Dim n1 As Integer
Dim n2 As Integer
n1 = Val(Text1.Text)
n2 = Val(Label11.Caption)
If n1 > n2 Then
MsgBox "��治�㣡�������������!" & vbCrLf & "���Ϊ:" & n2 & "����Ҫ����:" & n1 & "��", , "����"
Text1.SetFocus
End If
End Sub
