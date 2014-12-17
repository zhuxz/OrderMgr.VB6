VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "账单管理系统"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu menuMgr 
      Caption         =   "管理"
      Begin VB.Menu menuOrders 
         Caption         =   "服务单预览"
      End
      Begin VB.Menu menuEmployee 
         Caption         =   "员工"
      End
      Begin VB.Menu menuRoom 
         Caption         =   "房间"
      End
      Begin VB.Menu menuOrder 
         Caption         =   "服务单"
      End
   End
   Begin VB.Menu menuCofig 
      Caption         =   "配置"
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    Dim frmName As String
    frmName = m_ini.ReadString(mDefine.INISEC_MAIN, mDefine.INIKEY_ACTIVEFORM, "")
    Select Case frmName
        Case mDefine.FORMNAME_CONFIG: ShowForm_Config
        Case mDefine.FORMNAME_ORDERS: ShowForm_Orders
        Case mDefine.FORMNAME_EMPLOYEE: ShowForm_Employee
        Case mDefine.FORMNAME_ROOM: ShowForm_Room
        Case mDefine.FORMNAME_ORDER: ShowForm_Order
    End Select
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    m_db.Close
    Set m_db = Nothing
    
    Set m_ini = Nothing
End Sub

Private Sub menuCofig_Click()
    ShowForm_Config
End Sub

Private Sub menuEmployee_Click()
    ShowForm_Employee
End Sub

Private Sub menuOrder_Click()
    ShowForm_Order
End Sub

Private Sub menuOrders_Click()
    ShowForm_Orders
End Sub

Private Sub menuRoom_Click()
    ShowForm_Room
End Sub
