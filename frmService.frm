VERSION 5.00
Begin VB.Form frmService 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "添加服务项目"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4950
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
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtServicePrice 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddService 
      Caption         =   "添加"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtServiceName 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdCloseWin 
      Caption         =   "关闭"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "单价："
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "名称："
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   675
   End
End
Attribute VB_Name = "frmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_action As Long
Public m_data As Variant

Private Sub cmdAddService_Click()
    Dim Data As Variant
    Dim valid As String
    
    Data = GetService(valid)
    If Len(valid) = 0 Then
        SaveService Data
    Else
        MsgBox valid, vbExclamation, mDefine.MSG_TITLE
    End If
End Sub

Private Sub cmdCloseWin_Click()
    CloseAndRefreshWin
End Sub

Private Sub Form_Activate()
    If m_action = MgrAction.update_ Then
        PopulateService
    Else
        ResetService
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        CloseAndRefreshWin
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetService(ByRef Validation As String)
    Dim ret(Service.BOF_ + 1 To Service.EOF_ - 1) As Variant
    ret(Service.name_) = Trim$(Me.txtServiceName.Text)
    If Len(ret(Service.name_)) = 0 Then
        Validation = mDefine.MSG_VALIDSERVICENAME
    End If
    ret(Service.price) = Trim$(Me.txtServicePrice.Text)
    GetService = ret
End Function

Private Sub ResetService()
    Me.txtServiceName.Text = ""
    Me.txtServicePrice.Text = ""
End Sub

Private Sub PopulateService()
    If IsArray(m_data) Then
        Me.txtServiceName.Text = m_data(Service.name_)
        Me.txtServicePrice.Text = m_data(Service.price)
    End If
End Sub

Private Sub CloseWin()
    Me.Hide
End Sub

Private Sub CloseAndRefreshWin()
    Me.Hide
    m_frmServices.RefreshServices
End Sub

Private Sub SaveService(ByVal vData As Variant)
    Dim sql As String
    sql = "INSERT INTO " & mDefine.DBTN_SERVICES & "([desc], [price]) VALUES('" & vData(Service.name_) _
        & "', " & vData(Service.price) & ")"
    m_db.Execute sql
    If MsgBox(mDefine.MSG_ADDEMPLOYEEASK, vbYesNo, mDefine.MSG_TITLE) = vbYes Then
        ResetService
        Me.txtServiceName.SetFocus
    Else
        CloseAndRefreshWin
    End If
    'Dim rs As ADODB.Recordset
    'sql = "SELECT * FROM " & mDefine.DBTN_Service '& _
        " WHERE name='" & ServiceData(Service.name_) & "'"
    'rs.Open sql, m_db, adOpenDynamic, adLockOptimistic
    'Set rs = m_db.Execute(sql)
    'rs.AddNew
End Sub
