VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmEmployees 
   Caption         =   "员工"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11925
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
   ScaleHeight     =   7605
   ScaleWidth      =   11925
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdShowAllEmployees 
      Caption         =   "显示全部"
      Height          =   375
      Left            =   10560
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeleteEmployee 
      Caption         =   "删除"
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin VB.CheckBox ckQuery 
      Caption         =   "查询"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdAddEmployee 
      Caption         =   "添加"
      Height          =   375
      Index           =   1
      Left            =   5160
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvEmployees 
      Height          =   5895
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10398
      View            =   3
      LabelEdit       =   1
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
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton cmdSearchEmployees 
         Caption         =   "确定"
         Height          =   360
         Left            =   3600
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cboEmployeeSex 
         Height          =   360
         ItemData        =   "frmEmployees.frx":0000
         Left            =   840
         List            =   "frmEmployees.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtEmployeeName 
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "性别："
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "姓名："
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ckQuery_Click()
    EnableQuery CBool(Me.ckQuery.value)
    If Me.ckQuery.value = 0 Then
        PopulateEmployees
    End If
End Sub

Private Sub cmdAddEmployee_Click(Index As Integer)
    If m_frmEmployee Is Nothing Then Set m_frmEmployee = New frmEmployee
    m_frmEmployee.m_action = MgrAction.add_
    m_frmEmployee.Show 1
End Sub

Private Sub cmdDeleteEmployee_Click()
    Dim lvItem As ListItem
    Dim del_ids As Variant
    For Each lvItem In Me.lvEmployees.ListItems
        If lvItem.Checked Then
            AppendToVariantArr del_ids, Mid(lvItem.key, 4)
        End If
    Next
    If IsArray(del_ids) Then
        DeleteEmployeesByIds del_ids
        RefreshEmployees
    End If
End Sub

Private Sub cmdSearchEmployees_Click()
    Dim eplData As Variant
    eplData = GetEmployee()
    PopulateEmployees eplData
End Sub

Private Sub cmdShowAllEmployees_Click()
    PopulateEmployees
End Sub

Private Sub Form_Load()
    With Me.cboEmployeeSex
        .AddItem ""
        .AddItem mDefine.SEX_FEMALE
        .AddItem mDefine.SEX_MALE
        .Text = mDefine.SEX_FEMALE
    End With
    
    With Me.lvEmployees
        .ColumnHeaders.Add 1, "name", L_(LBL.name_), 1500
        .ColumnHeaders.Add 2, "sex", L_(LBL.sex), 1000
    End With
    
    EnableQuery False
    
    PopulateEmployees
End Sub

Private Sub lvEmployees_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lvEmployees
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
Private Function GetEmployee()
    Dim ret(Employee.BOF_ + 1 To Employee.EOF_ - 1) As Variant
    ret(Employee.name_) = Trim$(Me.txtEmployeeName.Text)
    ret(Employee.sex) = Me.cboEmployeeSex.Text
    GetEmployee = ret
End Function

Private Sub PopulateEmployees(Optional ByVal Filter As Variant = Empty)
    Dim Employees As Variant
    Dim i As Long
    Dim lvItem As ListItem
    Employees = LoadEmployeesFromDB(m_db, Filter)
    With Me.lvEmployees
        .ListItems.Clear
        If Not IsEmpty(Employees) Then
            For i = LBound(Employees) To UBound(Employees)
                Set lvItem = .ListItems.Add(i + 1, "id " & (Employees(i)(Employee.ID)), Employees(i)(Employee.name_))
                lvItem.SubItems(1) = Employees(i)(Employee.sex)
            Next
        End If
    End With
End Sub

Private Sub EnableQuery(ByVal Enabled As Boolean)
    Me.txtEmployeeName.Enabled = Enabled
    Me.cboEmployeeSex.Enabled = Enabled
    Me.cmdSearchEmployees.Enabled = Enabled
End Sub

Public Sub RefreshEmployees()
    If Me.ckQuery.value Then
        cmdSearchEmployees_Click
    Else
        PopulateEmployees
    End If
End Sub
