VERSION 5.00
Begin VB.Form frmEmployee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "添加员工"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5100
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
   ScaleHeight     =   1530
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCloseWin 
      Caption         =   "关闭"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtEmployeeName 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.ComboBox cboEmployeeSex 
      Height          =   360
      ItemData        =   "frmEmployee.frx":0000
      Left            =   960
      List            =   "frmEmployee.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddEmployee 
      Caption         =   "添加"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "姓名："
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "性别："
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   675
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_action As Long
Public m_data As Variant

Private Sub cmdAddEmployee_Click()
    Dim eplData As Variant
    Dim valid As String
    
    eplData = GetEmployee(valid)
    
    If Len(valid) = 0 Then
        SaveEmployee eplData
    Else
        MsgBox valid, vbExclamation, mDefine.MSG_TITLE
    End If
End Sub

Private Sub cmdCloseWin_Click()
    CloseAndRefreshWin
End Sub

Private Sub Form_Activate()
    If m_action = MgrAction.update_ Then
        PopulateEmployee
    Else
        ResetEmployee
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        CloseAndRefreshWin
    End If
End Sub

Private Sub Form_Load()
    With Me.cboEmployeeSex
        .AddItem mDefine.SEX_FEMALE
        .AddItem mDefine.SEX_MALE
        .ListIndex = 0
    End With
    
    If m_action = MgrAction.update_ Then
        PopulateEmployee
    Else
        ResetEmployee
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetEmployee(ByRef Validation As String)
    Dim ret(Employee.BOF_ + 1 To Employee.EOF_ - 1) As Variant
    ret(Employee.name_) = Trim$(Me.txtEmployeeName.Text)
    If Len(ret(Employee.name_)) = 0 Then
        Validation = mDefine.MSG_VALIDEMPLOYEENAME
    End If
    ret(Employee.sex) = Trim$(Me.cboEmployeeSex.Text)
    GetEmployee = ret
End Function

Private Sub SaveEmployee(ByVal vData As Variant)
    Dim sql As String
    sql = "INSERT INTO " & mDefine.DBTN_EMPLOYEES & "([name], [sex]) VALUES('" & vData(Employee.name_) _
        & "', '" & vData(Employee.sex) & "')"
    m_db.Execute sql
    If MsgBox(mDefine.MSG_ADDEMPLOYEEASK, vbYesNo, mDefine.MSG_TITLE) = vbYes Then
        ResetEmployee
    Else
        CloseAndRefreshWin
    End If
    'Dim rs As ADODB.Recordset
    'sql = "SELECT * FROM " & mDefine.DBTN_EMPLOYEE '& _
        " WHERE name='" & EmployeeData(Employee.name_) & "'"
    'rs.Open sql, m_db, adOpenDynamic, adLockOptimistic
    'Set rs = m_db.Execute(sql)
    'rs.AddNew
End Sub

Private Sub ResetEmployee()
    Me.txtEmployeeName.Text = ""
End Sub

Private Sub PopulateEmployee()
    If IsArray(m_data) Then
        Me.txtEmployeeName.Text = m_data(Employee.name_)
        Me.cboEmployeeSex.ListIndex = m_data(Employee.sex)
    End If
End Sub

Private Sub CloseWin()
    Me.Hide
End Sub

Private Sub CloseAndRefreshWin()
    Me.Hide
    m_frmEmployees.RefreshEmployees
End Sub

