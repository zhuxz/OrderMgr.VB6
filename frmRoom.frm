VERSION 5.00
Begin VB.Form frmRoom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "添加房间"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
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
   ScaleHeight     =   1080
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAddRoom 
      Caption         =   "添加"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtRoomName 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdCloseWin 
      Caption         =   "关闭"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "房号："
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   675
   End
End
Attribute VB_Name = "frmRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_action As Long
Public m_data As Variant

Private Sub cmdAddRoom_Click()
    Dim Data As Variant
    Dim valid As String
    
    Data = GetRoom(valid)
    If Len(valid) = 0 Then
        SaveRoom Data
    Else
        MsgBox valid, vbExclamation, mDefine.MSG_TITLE
    End If
End Sub

Private Sub cmdCloseWin_Click()
    CloseAndRefreshWin
End Sub

Private Sub Form_Activate()
    If m_action = MgrAction.update_ Then
        PopulateRoom
    Else
        ResetRoom
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
Private Function GetRoom(ByRef Validation As String)
    Dim ret(Room.BOF_ + 1 To Room.EOF_ - 1) As Variant
    ret(Room.name_) = Trim$(Me.txtRoomName.Text)
    If Len(ret(Room.name_)) = 0 Then
        Validation = MSG_(MSG.ValidRoomName)
    End If
    GetRoom = ret
End Function

Private Sub ResetRoom()
    Me.txtRoomName.Text = ""
End Sub

Private Sub PopulateRoom()
    If IsArray(m_data) Then
        Me.txtRoomName.Text = m_data(Room.name_)
    End If
End Sub

Private Sub CloseWin()
    Me.Hide
End Sub

Private Sub CloseAndRefreshWin()
    Me.Hide
    m_frmRooms.RefreshRooms
End Sub

Private Sub SaveRoom(ByVal vData As Variant)
    Dim sql As String
    sql = "INSERT INTO " & TBN_(TBN.Rooms) & "([name]) VALUES('" & vData(Room.name_) _
        & "')"
    m_db.Execute sql
    If MsgBox(mDefine.MSG_ADDEMPLOYEEASK, vbYesNo, mDefine.MSG_TITLE) = vbYes Then
        ResetRoom
    Else
        CloseAndRefreshWin
    End If
    'Dim rs As ADODB.Recordset
    'sql = "SELECT * FROM " & mDefine.DBTN_Room '& _
        " WHERE name='" & RoomData(Room.name_) & "'"
    'rs.Open sql, m_db, adOpenDynamic, adLockOptimistic
    'Set rs = m_db.Execute(sql)
    'rs.AddNew
End Sub


