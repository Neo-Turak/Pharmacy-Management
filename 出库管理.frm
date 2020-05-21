VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form 出库管理 
   Caption         =   "出库管理"
   ClientHeight    =   8820
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10935
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
   ScaleHeight     =   8820
   ScaleWidth      =   10935
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10815
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "出库管理.frx":0000
         Left            =   5760
         List            =   "出库管理.frx":0025
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "删除0库存"
         Height          =   495
         Left            =   9480
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "排序"
         Height          =   495
         Left            =   7440
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "打印"
         Height          =   495
         Left            =   4320
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "盘点"
         Height          =   495
         Left            =   3000
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "出库管理.frx":0079
         Left            =   120
         List            =   "出库管理.frx":008C
         TabIndex        =   3
         Text            =   "全部"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "查询"
         Height          =   495
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4080
      Top             =   8280
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      RecordSource    =   "出库"
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
      Bindings        =   "出库管理.frx":00B0
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   5741
      _Version        =   393216
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
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
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
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
         DataField       =   "出库单号"
         Caption         =   "出库单号"
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
         DataField       =   "取药单位"
         Caption         =   "取药单位"
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
         DataField       =   "药品名"
         Caption         =   "药品名"
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
      BeginProperty Column07 
         DataField       =   "售价"
         Caption         =   "售价"
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
      BeginProperty Column08 
         DataField       =   "总价"
         Caption         =   "总价"
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
      BeginProperty Column09 
         DataField       =   "出口时间"
         Caption         =   "出口时间"
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
         DataField       =   "有效期"
         Caption         =   "有效期"
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
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "出库管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo err1:

Dim Conn As ADODB.Connection
Dim rt As ADODB.Recordset
Dim co As String, sql As String
Set Conn = New ADODB.Connection
Set rt = New ADODB.Recordset
co = "provider=microsoft.jet.OLEDB.4.0;DATA SOURCE=" & App.Path & "\logs.MDB;Persist security info=false;"

If Combo1.Text = "全部" Then
sql = "select * from 出库"
Else
sql = "select * from 出库 where 取药单位='" & Combo1.Text & "'"
End If

Conn.Open co
Conn.CursorLocation = adUseClient
rt.Open sql, Conn, 3, 2
Set Adodc1.Recordset = rt
Set DataGrid1.DataSource = rt
Exit Sub
err1:
MsgBox "出现错误，我也没办法！" & vbCrLf & "错误编号：" & Err.Number & vbCrLf & "错误描述：" & Err.Description
End Sub

Private Sub Command2_Click()
 Dim count As Integer
 Dim money As Currency
 count = 0
 money = 0
If Adodc1.Recordset.RecordCount <> 0 Then
Adodc1.Recordset.MoveFirst
Do Until Adodc1.Recordset.EOF = True
 count = count + Adodc1.Recordset.Fields("数量").Value
 money = money + Adodc1.Recordset.Fields("总价").Value
Adodc1.Recordset.MoveNext
Loop
MsgBox "当前条目：" & Adodc1.Recordset.RecordCount & vbCrLf & _
"药品数量：" & count & vbCrLf & _
"总价：" & money & "元" & vbCrLf & _
"不重复药品数量：" & Repeat() & vbCrLf & _
"重复药品类：" & Repeater(), , "结果"
End If
End Sub

Private Sub Command3_Click()
Printer.PaperSize = 9
'开始打印段落
'进行对话框属性设置
Printer.Orientation = 1
Printer.ScaleWidth = 21
Printer.ScaleHeight = 27
Printer.FontBold = True
Printer.FontItalic = True
Printer.FontSize = 16
Printer.CurrentX = 10
Printer.CurrentY = 2
Printer.Print "出库清单"      'Title

Printer.FontBold = False
Printer.FontItalic = False
Printer.FontSize = 12
 Adodc1.Recordset.MoveFirst
 Dim rt As ADODB.Recordset
 Set rt = Adodc1.Recordset
 With rt
 n = 4
 Printer.CurrentX = 1
 Printer.CurrentY = 3
 Printer.FontBold = True
 Printer.Print "ID" & Space(2) & "单号" & Space(2) & "单位" & Space(2) & "药品名" & Space(5) & "数量" & Space(2) & "剂型" & Space(2) & "规格" & Space(2) & "售价" & Space(2) & "总价" & "出库时间" & Space(2) & _
 "有效期至"
 Printer.FontBold = False
Do Until rt.EOF = True

Printer.CurrentX = 1
Printer.CurrentY = n

 Printer.Print .Fields(0).Value & Space(2) & .Fields(1).Value & Space(2) & .Fields(2).Value & Space(2) & .Fields(3).Value & Space(2) & .Fields(4).Value & Space(2) & .Fields(5).Value & _
   Space(2) & .Fields(6).Value & Space(2) & .Fields(7).Value & Space(2) & .Fields(8).Value & Space(2) & .Fields(9).Value
 rt.MoveNext
 n = n + 1
 Loop
 Printer.Print vbTab & "----------------------------------------------------------------------"

 Printer.EndDoc
 
End With
End Sub

Private Sub Command4_Click()
Dim Conn As ADODB.Connection
Dim rt As ADODB.Recordset
Dim co As String, sql As String

Set Conn = New ADODB.Connection
Set rt = New ADODB.Recordset
co = "provider=microsoft.jet.OLEDB.4.0;DATA SOURCE=" & App.Path & "\logs.MDB;Persist security info=false;"
If Combo1.Text = "全部" Then
sql = "select * from 出库 order by " & Combo2.Text
Else
sql = "select * from 出库 where 取药单位='" & Combo1.Text & "' order by " & Combo2.Text
End If
Conn.Open co
Conn.CursorLocation = adUseClient
rt.Open sql, Conn, 3, 2
Set Adodc1.Recordset = rt
Set DataGrid1.DataSource = rt
End Sub

Private Sub Command5_Click()
On Error GoTo err1:
If MsgBox("删除0库存？ 不存在的，别开玩笑！ 继续执行吗？", 64 + vbOKCancel, "警告") = vbOK Then
 Dim Conn As ADODB.Connection
 Dim rs As ADODB.Recordset
 Dim sql As String
 Set Conn = New ADODB.Connection
 Set rs = New ADODB.Recordset
 sql = "delete  from 出库 where 数量=0"
 Conn.Open "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\logs.mdb;persist security info=false"
 Conn.CursorLocation = adUseClient
 rs.Open sql, Conn, 3, 2
 MsgBox "删除0库存记录成功！"
 Exit Sub
err1:
 MsgBox "出现错误！" & vbcrl & Err.Number & vbCrLf & "错误描述：" & Err.Description
End If

End Sub

Private Sub Form_Load()
DataGrid1.Width = Screen.Width
DataGrid1.Height = Screen.Height

End Sub

Public Function Repeat() As Integer
Dim Conn As ADODB.Connection
Dim rt As ADODB.Recordset
Dim co As String, sql As String

Set Conn = New ADODB.Connection
Set rt = New ADODB.Recordset
co = "provider=microsoft.jet.OLEDB.4.0;DATA SOURCE=" & App.Path & "\logs.MDB;Persist security info=false;"
sql = "select distinct * from 出库"
Conn.Open co
Conn.CursorLocation = adUseClient
rt.Open sql, Conn, 3, 2
Repeat = rt.RecordCount
End Function

Public Function Repeater() As Integer
Dim Conn As ADODB.Connection
Dim rt As ADODB.Recordset
Dim co As String, sql As String

Set Conn = New ADODB.Connection
Set rt = New ADODB.Recordset
co = "provider=microsoft.jet.OLEDB.4.0;DATA SOURCE=" & App.Path & "\logs.MDB;Persist security info=false;"
sql = "Select 药品名,规格,Count(*) From 出库 Group By 药品名,规格 Having Count(*) > 1"
Conn.Open co
Conn.CursorLocation = adUseClient
rt.Open sql, Conn, 3, 2
Repeater = rt.RecordCount
End Function
