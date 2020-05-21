VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm 库房管理系统 
   BackColor       =   &H80000002&
   Caption         =   "库房工作站"
   ClientHeight    =   9495
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13170
   Icon            =   "库房管理系统.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   9000
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2017-04-10"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "日期"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "6:40"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "时间"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "用户名"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "部门"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "职位"
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7911
            Text            =   "荒地镇卫生院"
            TextSave        =   "荒地镇卫生院"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "医院名称"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu 日常工作 
      Caption         =   "日常工作(&F)"
      Index           =   1
      Begin VB.Menu 药品入库 
         Caption         =   "药品入库"
         Index           =   11
         Shortcut        =   {F1}
      End
      Begin VB.Menu 药品出口 
         Caption         =   "药品出库"
         Index           =   12
         Shortcut        =   {F2}
      End
      Begin VB.Menu 库存管理 
         Caption         =   "库存管理"
         Index           =   13
         Shortcut        =   {F3}
      End
      Begin VB.Menu 出库管理 
         Caption         =   "出库管理"
         Shortcut        =   {F4}
      End
      Begin VB.Menu 药品报损 
         Caption         =   "药品报损"
         Shortcut        =   {F5}
      End
      Begin VB.Menu 药库盘点 
         Caption         =   "药库盘点"
         Index           =   13
      End
   End
   Begin VB.Menu 资料管理 
      Caption         =   "资料管理(&R)"
      Index           =   2
      Begin VB.Menu 药典设置 
         Caption         =   "药典设置"
         Shortcut        =   {F7}
      End
      Begin VB.Menu 信息提示 
         Caption         =   "信息提示器"
         Index           =   4
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu 收支管理 
      Caption         =   "收支管理"
      Begin VB.Menu 总收支 
         Caption         =   "总收支"
      End
      Begin VB.Menu 各部门收支 
         Caption         =   "各部门收支"
      End
   End
   Begin VB.Menu 设置 
      Caption         =   "设置"
      Begin VB.Menu 供药单位设置 
         Caption         =   "供药单位设置"
      End
      Begin VB.Menu 取药单位设置 
         Caption         =   "取药单位设置"
      End
      Begin VB.Menu 报警器 
         Caption         =   "报警器"
      End
      Begin VB.Menu 密码修改 
         Caption         =   "密码修改"
      End
   End
End
Attribute VB_Name = "库房管理系统"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub 报警器_Click()
药库工作站.设置.Show
End Sub

Private Sub 出库管理_Click()
药库工作站.出库管理.Show
End Sub

Private Sub 各部门收支_Click()
MsgBox "预留功能，后期添加！"
End Sub

Private Sub 供药单位设置_Click()
MsgBox "预留功能，后期添加！"
End Sub

Private Sub 库存管理_Click(Index As Integer)
药品库存.Show
End Sub

Private Sub 密码修改_Click()
MsgBox "预留功能，后期添加！"
End Sub

Private Sub 取药单位设置_Click()
MsgBox "预留功能，后期添加！"
End Sub

Private Sub 信息添加_Click(Index As Integer)
MsgBox "预留功能，后期添加！"
End Sub

Private Sub 信息提示器_Click(Index As Integer)

End Sub

Private Sub 信息修改_Click(Index As Integer)
MsgBox "预留功能，后期添加！"
End Sub

Private Sub 信息提示_Click(Index As Integer)
信息提示器.Show
End Sub

Private Sub 药典设置_Click()
药库工作站.药典设置.Show
End Sub

Private Sub 药品报废申请_Click(Index As Integer)
报废申请.Show
End Sub

Private Sub 药库盘点_Click(Index As Integer)
MsgBox "预留功能，后期添加！"
End Sub

Private Sub 药品报损_Click()
MsgBox "预留功能，后期添加！"
End Sub

Private Sub 药品出口_Click(Index As Integer)
药品出库.Show
End Sub

Private Sub 药品入库_Click(Index As Integer)
药品进库.Show
End Sub

Private Sub 总收支_Click()
MsgBox "预留功能，后期添加！"
End Sub
