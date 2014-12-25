VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "账单管理系统"
   ClientHeight    =   7695
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11940
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu menuMgr 
      Caption         =   "管理"
      Begin VB.Menu menuOrders 
         Caption         =   "服务单预览"
      End
      Begin VB.Menu menuEmployee 
         Caption         =   "员工"
      End
      Begin VB.Menu menuRooms 
         Caption         =   "房间"
      End
      Begin VB.Menu menuServices 
         Caption         =   "服务项目"
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
    Me.Height = 8400
    Me.Width = 12100
    
    Dim frmName As String
    frmName = m_ini.ReadString(mDefine.INISEC_MAIN, mDefine.INIKEY_ACTIVEFORM, "")
    Select Case frmName
        Case mDefine.FORMNAME_CONFIG: ShowForm_Config
        Case mDefine.FORMNAME_ORDERS: ShowForm_Orders
        Case mDefine.FORMNAME_EMPLOYEES: ShowForm_Employees
        Case FN_(FRMN.Rooms): ShowForm_Rooms
        Case mDefine.FORMNAME_ORDER: ShowForm_Order
        Case mDefine.FORMNAME_SERVICES: ShowForm_Services
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
    ShowForm_Employees
End Sub

Private Sub menuOrder_Click()
    ShowForm_Order
End Sub

Private Sub menuOrders_Click()
    ShowForm_Orders
End Sub

Private Sub menuRooms_Click()
    ShowForm_Rooms
End Sub

Private Sub menuServices_Click()
    ShowForm_Services
End Sub
