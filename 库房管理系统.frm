VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm �ⷿ����ϵͳ 
   BackColor       =   &H80000002&
   Caption         =   "�ⷿ����վ"
   ClientHeight    =   9495
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13170
   Icon            =   "�ⷿ����ϵͳ.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '����ȱʡ
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
            Object.ToolTipText     =   "����"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "6:40"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "ʱ��"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "�û���"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "����"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "ְλ"
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7911
            Text            =   "�ĵ�������Ժ"
            TextSave        =   "�ĵ�������Ժ"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "ҽԺ����"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu �ճ����� 
      Caption         =   "�ճ�����(&F)"
      Index           =   1
      Begin VB.Menu ҩƷ��� 
         Caption         =   "ҩƷ���"
         Index           =   11
         Shortcut        =   {F1}
      End
      Begin VB.Menu ҩƷ���� 
         Caption         =   "ҩƷ����"
         Index           =   12
         Shortcut        =   {F2}
      End
      Begin VB.Menu ������ 
         Caption         =   "������"
         Index           =   13
         Shortcut        =   {F3}
      End
      Begin VB.Menu ������� 
         Caption         =   "�������"
         Shortcut        =   {F4}
      End
      Begin VB.Menu ҩƷ���� 
         Caption         =   "ҩƷ����"
         Shortcut        =   {F5}
      End
      Begin VB.Menu ҩ���̵� 
         Caption         =   "ҩ���̵�"
         Index           =   13
      End
   End
   Begin VB.Menu ���Ϲ��� 
      Caption         =   "���Ϲ���(&R)"
      Index           =   2
      Begin VB.Menu ҩ������ 
         Caption         =   "ҩ������"
         Shortcut        =   {F7}
      End
      Begin VB.Menu ��Ϣ��ʾ 
         Caption         =   "��Ϣ��ʾ��"
         Index           =   4
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu ��֧���� 
      Caption         =   "��֧����"
      Begin VB.Menu ����֧ 
         Caption         =   "����֧"
      End
      Begin VB.Menu ��������֧ 
         Caption         =   "��������֧"
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu ��ҩ��λ���� 
         Caption         =   "��ҩ��λ����"
      End
      Begin VB.Menu ȡҩ��λ���� 
         Caption         =   "ȡҩ��λ����"
      End
      Begin VB.Menu ������ 
         Caption         =   "������"
      End
      Begin VB.Menu �����޸� 
         Caption         =   "�����޸�"
      End
   End
End
Attribute VB_Name = "�ⷿ����ϵͳ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ������_Click()
ҩ�⹤��վ.����.Show
End Sub

Private Sub �������_Click()
ҩ�⹤��վ.�������.Show
End Sub

Private Sub ��������֧_Click()
MsgBox "Ԥ�����ܣ�������ӣ�"
End Sub

Private Sub ��ҩ��λ����_Click()
MsgBox "Ԥ�����ܣ�������ӣ�"
End Sub

Private Sub ������_Click(Index As Integer)
ҩƷ���.Show
End Sub

Private Sub �����޸�_Click()
MsgBox "Ԥ�����ܣ�������ӣ�"
End Sub

Private Sub ȡҩ��λ����_Click()
MsgBox "Ԥ�����ܣ�������ӣ�"
End Sub

Private Sub ��Ϣ���_Click(Index As Integer)
MsgBox "Ԥ�����ܣ�������ӣ�"
End Sub

Private Sub ��Ϣ��ʾ��_Click(Index As Integer)

End Sub

Private Sub ��Ϣ�޸�_Click(Index As Integer)
MsgBox "Ԥ�����ܣ�������ӣ�"
End Sub

Private Sub ��Ϣ��ʾ_Click(Index As Integer)
��Ϣ��ʾ��.Show
End Sub

Private Sub ҩ������_Click()
ҩ�⹤��վ.ҩ������.Show
End Sub

Private Sub ҩƷ��������_Click(Index As Integer)
��������.Show
End Sub

Private Sub ҩ���̵�_Click(Index As Integer)
MsgBox "Ԥ�����ܣ�������ӣ�"
End Sub

Private Sub ҩƷ����_Click()
MsgBox "Ԥ�����ܣ�������ӣ�"
End Sub

Private Sub ҩƷ����_Click(Index As Integer)
ҩƷ����.Show
End Sub

Private Sub ҩƷ���_Click(Index As Integer)
ҩƷ����.Show
End Sub

Private Sub ����֧_Click()
MsgBox "Ԥ�����ܣ�������ӣ�"
End Sub
