VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form 药品进库 
   BackColor       =   &H00E0E0E0&
   Caption         =   "药品入库单"
   ClientHeight    =   10050
   ClientLeft      =   2145
   ClientTop       =   1935
   ClientWidth     =   14700
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10050
   ScaleWidth      =   14700
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "保存"
      Height          =   495
      Left            =   6360
      TabIndex        =   35
      Top             =   4200
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   7440
      TabIndex        =   27
      Top             =   3720
      Width           =   6495
      Begin VB.CommandButton Command4 
         Caption         =   "总"
         Height          =   375
         Left            =   5760
         TabIndex        =   30
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """￥""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4680
         TabIndex        =   33
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label16 
         BackColor       =   &H0000C000&
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1080
         TabIndex        =   31
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """￥""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2520
         TabIndex        =   29
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "总数量："
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "总价："
         Height          =   375
         Left            =   1920
         TabIndex        =   32
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "预期利润："
         Height          =   375
         Left            =   3600
         TabIndex        =   34
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查询"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      ItemData        =   "药品进库单.frx":0000
      Left            =   1200
      List            =   "药品进库单.frx":0019
      TabIndex        =   1
      Text            =   "速记码"
      Top             =   1560
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "药品进库单.frx":0055
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   1931
      _Version        =   393216
      BackColor       =   32768
      ForeColor       =   -2147483634
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      BeginProperty Column06 
         DataField       =   "采购价"
         Caption         =   "采购价"
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
            Alignment       =   2
            DividerStyle    =   4
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2264.882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3119.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2099.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   854.929
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   11160
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "添加"
      Height          =   495
      Left            =   6360
      TabIndex        =   9
      Top             =   3720
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "药品进库单.frx":006A
      Height          =   4575
      Left            =   120
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4800
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   8070
      _Version        =   393216
      BackColor       =   -2147483627
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "入库单号"
         Caption         =   "入库单号"
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
         DataField       =   "名称"
         Caption         =   "名称"
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "生产商"
         Caption         =   "生产商"
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
         DataField       =   "进价"
         Caption         =   "进价"
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
         DataField       =   "售价"
         Caption         =   "售价"
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
         DataField       =   "数量"
         Caption         =   "数量"
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
         DataField       =   "总价"
         Caption         =   "总价"
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
         DataField       =   "生产日期"
         Caption         =   "生产日期"
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
         DataField       =   "有效期至"
         Caption         =   "有效期至"
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
         DataField       =   "入库日期"
         Caption         =   "入库日期"
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
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1484.787
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1365.165
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5760
      Top             =   9600
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
      CommandType     =   1
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
      RecordSource    =   "select 入库单号,名称,剂型,规格,生产商,进价,售价,数量,总价,生产日期,有效期至,入库日期 from 入库 where 入库日期=date()"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   4320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   195493889
      CurrentDate     =   42458
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   4320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   195493889
      CurrentDate     =   42458
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """￥""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   2
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   26
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "总价："
      Height          =   375
      Left            =   4320
      TabIndex        =   25
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      DataField       =   "总代理商"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   4320
      TabIndex        =   24
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      DataField       =   "剂型"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   9720
      TabIndex        =   23
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      DataField       =   "规格"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   8280
      TabIndex        =   22
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      DataField       =   "通用名"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   1680
      TabIndex        =   21
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      DataField       =   "药品类型"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
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
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   11880
      TabIndex        =   19
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "关键字："
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "查询字段："
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1215
   End
   Begin MSForms.TextBox TextBox7 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   2055
      VariousPropertyBits=   746604571
      Size            =   "3625;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   1095
      ForeColor       =   0
      Caption         =   "入库单号"
      Size            =   "1931;661"
      BorderColor     =   255
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox6 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   3720
      Width           =   975
      VariousPropertyBits=   746604571
      Size            =   "1720;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Index           =   11
      Left            =   1680
      TabIndex        =   14
      Top             =   3720
      Width           =   615
      Caption         =   "售价"
      Size            =   "1085;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Index           =   8
      Left            =   3000
      TabIndex        =   13
      Top             =   4320
      Width           =   1215
      Caption         =   "有效日期至"
      Size            =   "2143;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   975
      Caption         =   "生产日期"
      Size            =   "1720;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox2 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   3720
      Width           =   855
      VariousPropertyBits=   746604571
      Size            =   "1508;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   615
      ForeColor       =   0
      Caption         =   "数量"
      Size            =   "1085;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox1 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
      VariousPropertyBits=   746604571
      Size            =   "3413;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   495
      Left            =   5400
      TabIndex        =   10
      Top             =   480
      Width           =   2535
      Caption         =   "药品入库表"
      Size            =   "4471;873"
      FontName        =   "叶根友毛笔行书2.0版"
      FontHeight      =   480
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "药品进库"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

On Error GoTo err1:
Dim ts As String
Dim FileName As String
Dim conn As ADODB.connection
Dim rs As ADODB.Recordset
Dim sql As String
Set conn = New ADODB.connection
Set rs = New ADODB.Recordset

Dim DbPw As String
FileName = App.Path & "\43.mdb"
ts = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbase\43.mdb;Persist Security Info=False"
sql = "select * from 药典 where " & Combo1.Text & " like '" & TextBox1.Text & "%'"
conn.Open ts
conn.CursorLocation = adUseClient
rs.Open sql, conn, adOpenKeyset, adLockOptimistic
Set Adodc2.Recordset = rs
Set DataGrid2.DataSource = rs
DataGrid2.Refresh
Exit Sub
err1:
    MsgBox "出错了！" & vbCrLf & "错误编号：" & Err.Number & " 错误描述：" & Err.Description
    Resume Next
End Sub
Private Sub Command2_Click()
On Error GoTo err1:
If TextBox7.Text <> "" And Label8.Caption <> "" And Label6.Caption <> "" Then
Dim rs As ADODB.Recordset
Set rs = Adodc1.Recordset
    On Error Resume Next
          rs.AddNew
          rs!入库单号 = TextBox7.Text
          rs!名称 = Label8.Caption
          rs!剂型 = Label10.Caption
          rs!规格 = Label9.Caption
          rs!生产商 = Label11.Caption
          rs!进价 = Label6.Caption
          rs!售价 = TextBox6.Text
          rs!数量 = TextBox2.Text
          rs!总价 = Label13.Caption
          rs!生产日期 = DTPicker1.Value
          rs!有效期至 = DTPicker2.Value
          rs!入库日期 = Date
           rs.update
           Else
           MsgBox "没填写必要内容！请核对后重试！", vbInformation + vbOKOnly + vbDefaultButton1, "错误警告"
           End If
      Exit Sub
err1:
      MsgBox "出现异常错误!" & vbCrLf & "错误编号：" & Err.Number & vbCrLf & "错误描述：" & Err.Description

End Sub

Private Sub Command3_Click()
On Error Resume Next
Adodc1.Recordset.update
End Sub

Private Sub TextBox5_Change()
On Error Resume Next
Label5.Caption = Trim(TextBox2.Text) * Trim(TextBox5.Text)
End Sub

Private Sub Command4_Click()
On Error GoTo err1:

Dim cnt As Integer
Dim Amont As Integer

Dim price As Currency
Dim Profit As Currency

cnt = Adodc1.Recordset.RecordCount

If cnt = 0 Then
Label15.Caption = 0
Else
Adodc1.Recordset.MoveFirst
Do Until Adodc1.Recordset.EOF = True
price = price + Adodc1.Recordset.Fields("总价").Value
Amont = Amont + Adodc1.Recordset.Fields("数量").Value
Profit = Profit + Adodc1.Recordset.Fields("售价").Value * Adodc1.Recordset.Fields("数量").Value
Adodc1.Recordset.MoveNext
Loop
Label15.Caption = price
Label16.Caption = Amont
Label18.Caption = Profit - price
End If
Exit Sub
err1:
MsgBox Err.Number
End Sub

Private Sub DataGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TextBox2.SetFocus
End If
End Sub

Private Sub TextBox1_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub


Private Sub TextBox2_LostFocus()
If TextBox2.Text <> "" Then
Label13.Caption = TextBox2.Text * Label6.Caption
End If

End Sub
