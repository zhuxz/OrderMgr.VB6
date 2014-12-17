VERSION 5.00
Begin VB.Form frmConfig 
   Caption         =   "系统配置"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12885
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
   ScaleHeight     =   7020
   ScaleWidth      =   12885
   WindowState     =   2  'Maximized
   Begin VB.Frame framDatabase 
      Caption         =   "数据库"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.CommandButton cmdDBTest 
         Caption         =   "测试连接"
         Height          =   375
         Left            =   9480
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdDBSave 
         Caption         =   "保存"
         Height          =   375
         Left            =   9480
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtDBAccount 
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   5760
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtDBServer 
         Height          =   405
         Left            =   1200
         TabIndex        =   5
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtDBPwd 
         Height          =   405
         Left            =   5760
         TabIndex        =   3
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtDBName 
         Height          =   405
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblDB 
         Caption         =   "密码："
         Height          =   255
         Left            =   4920
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblDBServer 
         AutoSize        =   -1  'True
         Caption         =   "服务器："
         Height          =   240
         Left            =   4920
         TabIndex        =   6
         Top             =   480
         Width           =   900
      End
      Begin VB.Label lblDBAccount 
         Caption         =   "帐号："
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblDBSource 
         AutoSize        =   -1  'True
         Caption         =   "数据库名："
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDBSave_Click()
    Dim dbSetting As mDefine.DatabaseSetting
    dbSetting = GetDbSetting()
    mMain.SaveDatabaseSetting dbSetting
End Sub

Private Sub cmdDBTest_Click()
    Dim dbString As String
    dbString = mMain.BuildDatabaseConnectString(Trim$(Me.txtDBName.Text), _
        Trim$(Me.txtDBServer.Text), Trim$(Me.txtDBAccount.Text), Trim$(Me.txtDBPwd.Text))

    Dim cn As ADODB.Connection
    Set cn = mMain.ConnectDatabase(dbString)
    If cn Is Nothing Then
        MsgBox mDefine.MSG_CONNECTDB_FAILED, vbExclamation, mDefine.MSG_TITLE
    Else
        MsgBox mDefine.MSG_CONNECTDB_SUCCESS, vbInformation, mDefine.MSG_TITLE
    End If
End Sub

Private Sub Form_Load()
    With m_dbSetting
        Me.txtDBServer.Text = .serverIP
        Me.txtDBName.Text = .source
        Me.txtDBAccount.Text = .account
        Me.txtDBPwd.Text = .pwd
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetDbSetting() As mDefine.DatabaseSetting
    Dim ret As mDefine.DatabaseSetting
    With ret
        .serverIP = Trim$(Me.txtDBServer.Text)
        .source = Trim$(Me.txtDBName.Text)
        .account = Trim$(Me.txtDBAccount.Text)
        .pwd = Trim$(Me.txtDBPwd.Text)
        .cnString = mMain.BuildDatabaseConnectString(.source, .serverIP, .account, .pwd)
    End With
    GetDbSetting = ret
End Function
