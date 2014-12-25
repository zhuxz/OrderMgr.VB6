VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmOrders 
   Caption         =   "服务单"
   ClientHeight    =   8475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   13110
   WindowState     =   2  'Maximized
   Begin VB.CheckBox ckQuery 
      Caption         =   "查询"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdShowAllOrders 
      Caption         =   "显示全部"
      Height          =   375
      Left            =   10560
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeleteOrders 
      Caption         =   "删除"
      Height          =   375
      Left            =   10080
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame framSearchEmployee 
      Caption         =   "查询"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   9855
      Begin VB.TextBox txtOrderPrice 
         Height          =   375
         Left            =   6480
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtRoomName 
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtEmployeeName 
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox cboEmployeeSex 
         Height          =   360
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdSearchOrders 
         Caption         =   "确定"
         Height          =   360
         Left            =   8520
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtServiceName 
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "单价："
         Height          =   240
         Left            =   5760
         TabIndex        =   16
         Top             =   480
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "房号："
         Height          =   240
         Left            =   2760
         TabIndex        =   15
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "服务项目："
         Height          =   240
         Left            =   2760
         TabIndex        =   14
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "性别："
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "姓名："
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdAddOrder 
      Cancel          =   -1  'True
      Caption         =   "添加"
      Height          =   360
      Left            =   10080
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvOrders 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11456
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ckQuery_Click()
    EnableQuery CBool(Me.ckQuery.value)
    If Me.ckQuery.value = 0 Then
        PopulateOrders
    End If
End Sub

Private Sub cmdAddOrder_Click()
    If m_frmOrder Is Nothing Then Set m_frmOrder = New frmOrder
    m_frmOrder.m_action = MgrAction.add_
    m_frmOrder.Show 1
End Sub

Private Sub cmdDeleteOrders_Click()
    Dim lvItem As ListItem
    Dim del_ids As Variant
    For Each lvItem In Me.lvOrders.ListItems
        If lvItem.Checked Then
            AppendToVariantArr del_ids, Mid(lvItem.key, 4)
        End If
    Next
    If IsArray(del_ids) Then
        DeleteOrdersByIds del_ids
        RefreshOrders
    End If
End Sub

Private Sub cmdSearchOrders_Click()
    Dim eplData As Variant
    eplData = GetOrder()
    PopulateOrders eplData
End Sub

Private Sub cmdShowAllOrders_Click()
    PopulateOrders
End Sub

Private Sub Form_Load()
    With Me.cboEmployeeSex
        .AddItem ""
        .AddItem mDefine.SEX_FEMALE
        .AddItem mDefine.SEX_MALE
        .ListIndex = 0
    End With
    
    With Me.lvOrders
        .ColumnHeaders.Add 1, "name", L_(LBL.name_), 1500
        .ColumnHeaders.Add 2, "sex", L_(LBL.sex), 1000
        .ColumnHeaders.Add 3, "serviceName", L_(LBL.service_name), 2000
        .ColumnHeaders.Add 4, "roomName", L_(LBL.room_name), 1000
        .ColumnHeaders.Add 5, "price", L_(LBL.price), 1500
        .ColumnHeaders.Add 6, "createDate", L_(LBL.createDate), 1500
        .ColumnHeaders.Add 7, "memo", L_(LBL.memo_), 1000
    End With
    
    EnableQuery False
    
    PopulateOrders
End Sub

Private Sub lvOrders_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lvOrders
        If (ColumnHeader.Index - 1) = .SortKey Then
            .SortOrder = (.SortOrder + 1) Mod 2
        Else
            .Sorted = False
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
        End If
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetOrder()
    Dim ret(Order.BOF_ + 1 To Order.EOF_ - 1) As Variant
    ret(Order.employeeName) = Trim$(Trim$(Me.txtEmployeeName.Text))
    ret(Order.serviceName) = Trim$(Trim$(Me.txtServiceName.Text))
    ret(Order.employeeSex) = Me.cboEmployeeSex.Text
    ret(Order.roomName) = Trim$(Me.txtRoomName.Text)
    If IsNumeric(Trim$(Me.txtOrderPrice.Text)) Then
        ret(Order.price) = Val(Trim$(Me.txtOrderPrice.Text))
    End If
    GetOrder = ret
End Function

Private Sub PopulateOrders(Optional ByVal Filter As Variant = Empty)
    Dim Orders As Variant
    Dim i As Long
    Dim lvItem As ListItem
    Orders = LoadOrdersFromDB(m_db, Filter)
    With Me.lvOrders
        .ListItems.Clear
        If Not IsEmpty(Orders) Then
            For i = LBound(Orders) To UBound(Orders)
                Set lvItem = .ListItems.Add(i + 1, "id " & (Orders(i)(Order.ID)), Orders(i)(Order.employeeName))
                lvItem.SubItems(1) = Orders(i)(Order.employeeSex)
                lvItem.SubItems(2) = Orders(i)(Order.serviceName)
                lvItem.SubItems(3) = Orders(i)(Order.roomName)
                lvItem.SubItems(4) = Orders(i)(Order.price)
                lvItem.SubItems(5) = Format$(Orders(i)(Order.createDate), "yyyy-mm-dd")
                lvItem.SubItems(6) = Orders(i)(Order.memo_)
            Next
        End If
    End With
End Sub

Private Sub EnableQuery(ByVal Enabled As Boolean)
    Me.txtServiceName.Enabled = Enabled
    Me.cboEmployeeSex.Enabled = Enabled
    Me.cmdSearchOrders.Enabled = Enabled
    Me.txtEmployeeName.Enabled = Enabled
    Me.txtOrderPrice.Enabled = Enabled
    Me.txtRoomName.Enabled = Enabled
End Sub

Public Sub RefreshOrders()
    If Me.ckQuery.value Then
        cmdSearchOrders_Click
    Else
        PopulateOrders
    End If
End Sub
