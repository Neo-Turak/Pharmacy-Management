VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ���� 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����"
   ClientHeight    =   3600
   ClientLeft      =   5310
   ClientTop       =   4455
   ClientWidth     =   3375
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   3375
   Begin VB.Frame Frame1 
      Caption         =   "�������ã�"
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2655
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   600
         Top             =   2040
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   873
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
      Begin VB.CommandButton Command1 
         Caption         =   "����"
         Height          =   495
         Left            =   720
         TabIndex        =   5
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text2 
         DataField       =   "����"
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text1 
         DataField       =   "����"
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "�������ޣ�      ��"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "�������ޣ�"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function update(ByVal s1 As String, ByVal s2 As String)
Dim conn As ADODB.connection
Dim rs As ADODB.Recordset
Dim sql As String, connection As String
Set conn = New ADODB.connection
Set rs = New ADODB.Recordset
connection = "provider=microsoft.jet.oledb.4.0;data source=dbase\43.mdb;persist security info=false"
sql = "update ������ set ����=" & s2 & ",����=" & s1 & " where id=1"
conn.Open connection
conn.CursorLocation = adUseClient
rs.Open sql, conn, 3, 2
End Function

Private Sub Command1_Click()
 Call update(Text1.Text, Text2.Text)
End Sub

Private Sub Form_Load()
On Error GoTo err1:
Dim conn As ADODB.connection
Dim rs As ADODB.Recordset
Dim sql As String, connection As String
Set conn = New ADODB.connection
Set rs = New ADODB.Recordset
connection = "provider=microsoft.jet.oledb.4.0;data source=dbase\43.mdb;persist security info=false"
sql = "select * from ������"
conn.Open connection
conn.CursorLocation = adUseClient
rs.Open sql, conn, 3, 2
Set Adodc1.Recordset = rs
Set Text1.DataSource = rs
Set Text2.DataSource = rs
Set Text1.DataSource = Nothing
Set Text2.DataSource = Nothing
Exit Sub
err1:
MsgBox "���ִ���" & Err.Number & "��������:" & Err.Description
End Sub
