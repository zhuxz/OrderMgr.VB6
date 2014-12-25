VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "账单添加"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5430
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMemo 
      Height          =   1095
      Left            =   960
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1680
      Width           =   4335
   End
   Begin MSComCtl2.DTPicker dtCreate 
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   108789761
      CurrentDate     =   41997
   End
   Begin VB.TextBox txtPrice 
      Height          =   360
      Left            =   3600
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox cboRoomName 
      Height          =   360
      ItemData        =   "frmOrder.frx":0000
      Left            =   960
      List            =   "frmOrder.frx":0002
      OLEDropMode     =   1  'Manual
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.ComboBox cboServiceName 
      Height          =   360
      ItemData        =   "frmOrder.frx":0004
      Left            =   960
      List            =   "frmOrder.frx":0006
      OLEDropMode     =   1  'Manual
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdAddOrder 
      Caption         =   "添加"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   2880
      Width           =   975
   End
   Begin VB.ComboBox cboEmployeeName 
      Height          =   360
      ItemData        =   "frmOrder.frx":0008
      Left            =   960
      List            =   "frmOrder.frx":000A
      OLEDropMode     =   1  'Manual
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdCloseWin 
      Caption         =   "关闭"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "备注："
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "日期："
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   12
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "单价："
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   11
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "房号："
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "服务项目："
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "姓名："
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_action As Long
Public m_data As Variant

Private m_employees As Variant
Private m_services As Variant
Private m_rooms As Variant

Private Sub cboServiceName_Click()
    Dim idx As Long
    Dim serviceData As Variant
    
    Me.txtPrice = GetDefaultPrice()
End Sub

Private Sub cmdAddOrder_Click()
    Dim eplData As Variant
    Dim valid As String
    
    eplData = GetOrder(valid)
    
    If Len(valid) = 0 Then
        SaveOrder eplData
    Else
        MsgBox valid, vbExclamation, mDefine.MSG_TITLE
    End If
End Sub

Private Sub cmdCloseWin_Click()
    CloseAndRefreshWin
End Sub

Private Sub Form_Activate()
    If m_action = MgrAction.update_ Then
        PopulateOrder
    Else
        ResetOrder
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        CloseAndRefreshWin
    End If
End Sub

Private Sub Form_Load()
    If m_action = MgrAction.update_ Then
        PopulateOrder
    Else
        ResetOrder
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetOrder(ByRef Validation As String)
    Dim ret(Order.BOF_ + 1 To Order.EOF_ - 1) As Variant
    Dim idx As Long
    
    idx = Me.cboEmployeeName.ListIndex
    If idx < 0 Then
        Validation = mDefine.MSG_VALIDEMPLOYEENAME
        Exit Function
    End If
    ret(Order.employeeId) = m_employees(idx)(Employee.ID)
    
    idx = Me.cboServiceName.ListIndex
    If idx < 0 Then
        Validation = MSG_(MSG.ValidServiceName)
        Exit Function
    End If
    ret(Order.serviceId) = m_services(idx)(Service.ID)
    
    idx = Me.cboRoomName.ListIndex
    If idx < 0 Then
        Validation = MSG_(MSG.ValidRoomName)
        Exit Function
    End If
    ret(Order.roomId) = m_rooms(idx)(Room.ID)
    
    Dim price As String
    price = Me.txtPrice.Text
    If IsNumeric(price) Then
        ret(Order.price) = Val(price)
    Else
        Validation = MSG_(MSG.ValidServicePrice)
        Exit Function
    End If
    
    ret(Order.memo_) = Me.txtMemo.Text
    ret(Order.createDate) = dtCreate.value
    
    GetOrder = ret
End Function

Private Sub SaveOrder(ByVal vData As Variant)
    Dim sql As String
    Dim fields As Variant
    Dim vals As Variant
    
    AppendToVariantArr fields, "[employeeId]"
    AppendToVariantArr vals, vData(Order.employeeId)
    
    AppendToVariantArr fields, "[roomId]"
    AppendToVariantArr vals, vData(Order.roomId)
    
    AppendToVariantArr fields, "[serviceId]"
    AppendToVariantArr vals, vData(Order.serviceId)
    
    AppendToVariantArr fields, "[price]"
    AppendToVariantArr vals, vData(Order.price)
    
    AppendToVariantArr fields, "[createDate]"
    AppendToVariantArr vals, "'" & vData(Order.createDate) & "'"
    
    AppendToVariantArr fields, "[memo]"
    AppendToVariantArr vals, "'" & vData(Order.memo_) & "'"
    
    
    sql = "INSERT INTO " & TBN_(TBN.Orders) & _
        "(" & Join(fields, ",") & ")" & _
        " VALUES(" & Join(vals, ",") & ")"
    m_db.Execute sql
    If MsgBox(mDefine.MSG_ADDEMPLOYEEASK, vbYesNo, mDefine.MSG_TITLE) = vbYes Then
        'ResetOrder
        Me.cboEmployeeName.SetFocus
    Else
        CloseAndRefreshWin
    End If
    'Dim rs As ADODB.Recordset
    'sql = "SELECT * FROM " & mDefine.DBTN_Order '& _
        " WHERE name='" & OrderData(Order.name_) & "'"
    'rs.Open sql, m_db, adOpenDynamic, adLockOptimistic
    'Set rs = m_db.Execute(sql)
    'rs.AddNew
End Sub

Private Sub ResetOrder()
    m_employees = mEmployee.LoadEmployeesFromDB(m_db)
    InitEmployeeNameComboBox m_employees
    
    m_services = mService.LoadServicesFromDB(m_db)
    InitServiceNameComboBox m_services
    
    m_rooms = mRoom.LoadRoomsFromDB(m_db)
    InitRoomNameComboBox m_rooms
    
    Me.txtPrice.Text = GetDefaultPrice()
    Me.txtMemo.Text = ""
    Me.dtCreate.value = Date
End Sub

Private Sub PopulateOrder()
    If IsArray(m_data) Then
'        Me.txtPrice.Text = m_data(Order.price)
'        Me.cboEmployeeName.ListIndex = m_data(Order.employeeId)
    End If
End Sub

Private Sub CloseWin()
    Me.Hide
End Sub

Private Sub CloseAndRefreshWin()
    Me.Hide
    m_frmOrders.RefreshOrders
End Sub

Private Sub InitEmployeeNameComboBox(ByVal Employees As Variant)
    If IsEmpty(Employees) Then
        Me.cboEmployeeName.Clear
        Exit Sub
    End If
    
    Dim i As Long
    Dim n As Long
    
    n = UBound(Employees)
    With Me.cboEmployeeName
        .Clear
        For i = 0 To n
            .AddItem Employees(i)(Employee.name_), i
        Next
        .ListIndex = 0
    End With
End Sub

Private Sub InitServiceNameComboBox(ByVal Data As Variant)
    If IsEmpty(Data) Then
        Me.cboServiceName.Clear
        Exit Sub
    End If
    
    Dim i As Long
    Dim n As Long
    
    n = UBound(Data)
    With Me.cboServiceName
        .Clear
        For i = 0 To n
            .AddItem Data(i)(Service.name_), i
        Next
        .ListIndex = 0
    End With
End Sub

Private Sub InitRoomNameComboBox(ByVal Data As Variant)
    If IsEmpty(Data) Then
        Me.cboRoomName.Clear
        Exit Sub
    End If
    
    Dim i As Long
    Dim n As Long
    
    n = UBound(Data)
    With Me.cboRoomName
        .Clear
        For i = 0 To n
            .AddItem Data(i)(Room.name_), i
        Next
        .ListIndex = 0
    End With
End Sub

Private Function GetDefaultPrice()
    If IsArray(m_services) Then
        Dim idx As Long
        idx = Me.cboServiceName.ListIndex
        If idx > -1 Then
            GetDefaultPrice = m_services(idx)(Service.price)
        End If
    End If
End Function


