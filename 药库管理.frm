VERSION 5.00
Begin VB.Form 药品管理 
   Caption         =   "库房管理系统"
   ClientHeight    =   9525
   ClientLeft      =   2430
   ClientTop       =   2280
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9525
   ScaleWidth      =   10515
   Begin VB.Menu 药品管理 
      Caption         =   "药品管理"
      Begin VB.Menu 药品入库 
         Caption         =   "药品入库"
      End
      Begin VB.Menu 药品查看 
         Caption         =   "药品查看"
      End
      Begin VB.Menu 药品出口 
         Caption         =   "药品出口"
      End
   End
   Begin VB.Menu 资料管理 
      Caption         =   "资料管理"
      Begin VB.Menu 药品资料管理 
         Caption         =   "药品资料管理"
      End
      Begin VB.Menu 药品目录 
         Caption         =   "药品目录"
      End
   End
End
Attribute VB_Name = "药品管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub 药品资料录入_Click()

End Sub
