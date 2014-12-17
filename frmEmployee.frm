VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmEmployee 
   Caption         =   "员工"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11400
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
   ScaleHeight     =   7125
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame framAddEmployee 
      Caption         =   "添加员工"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10455
      Begin VB.CommandButton cmdAddEmployee 
         Caption         =   "确定"
         Height          =   375
         Left            =   3600
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cboEmployeeSex 
         Height          =   360
         ItemData        =   "frmEmployee.frx":0000
         Left            =   840
         List            =   "frmEmployee.frx":0002
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtEmployeeName 
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "性别："
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "姓名："
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   675
      End
   End
   Begin MSComctlLib.ListView lvEmployees 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   9975
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddEmployee_Click()
    Dim eplData As Variant
    Dim valid As Boolean
    
    valid = True
    eplData = GetEmployee(valid)
    
    If valid Then
        SaveEmployee eplData
    Else
        MsgBox mDefine.MSG_NAMEVALID, vbExclamation, mDefine.MSG_TITLE
    End If
End Sub

Private Sub Form_Load()
    With Me.cboEmployeeSex
        .AddItem mDefine.SEX_FEMALE, mDefine.SEX_FEMALEID
        .AddItem mDefine.SEX_MALE, mDefine.SEX_MALEID
        .ListIndex = mDefine.SEX_FEMALEID
    End With
    
    With Me.lvEmployees
        .ColumnHeaders.Add 1, "name", "姓名", 1500
        .ColumnHeaders.Add 2, "sex", "性别", 1000
    End With
    RefreshEmployees
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetEmployee(ByRef Validation As Boolean)
    Dim ret(Employee.BOF_ + 1 To Employee.EOF_ - 1) As Variant
    ret(Employee.name_) = Trim$(Me.txtEmployeeName.Text)
    If Len(ret(Employee.name_)) = 0 Then
        Validation = False
    End If
    ret(Employee.sex) = Me.cboEmployeeSex.ListIndex
    GetEmployee = ret
End Function

Private Sub SaveEmployee(ByVal vData As Variant)
    Dim sql As String
    sql = "INSERT INTO " & mDefine.DBTN_EMPLOYEES & "([name], [sex]) VALUES('" & vData(Employee.name_) _
        & "', " & vData(Employee.sex) & ")"
    m_db.Execute sql
    'Dim rs As ADODB.Recordset
    'sql = "SELECT * FROM " & mDefine.DBTN_EMPLOYEE '& _
        " WHERE name='" & EmployeeData(Employee.name_) & "'"
    'rs.Open sql, m_db, adOpenDynamic, adLockOptimistic
    'Set rs = m_db.Execute(sql)
    'rs.AddNew
End Sub

Private Sub RefreshEmployees()
    Dim employees As Variant
    Dim i As Long
    employees = LoadEmployeesFromDB(m_db)
    
    If Not IsEmpty(employees) Then
        With Me.lvEmployees
            .ListItems.Clear
            For i = LBound(employees) To UBound(employees)
                .ListItems.Add i + 1, "key " & i, employees(i)(Employee.name_)
                .ListItems.item(i + 1).SubItems(1) = employees(i)(Employee.sex)
            Next
        End With
    End If
End Sub
