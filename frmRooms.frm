VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmRooms 
   Caption         =   "房间"
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
   Begin VB.CheckBox ckQuery 
      Caption         =   "查询"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdShowAllRooms 
      Caption         =   "显示全部"
      Height          =   375
      Left            =   10560
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeleteRooms 
      Caption         =   "删除"
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddRoom 
      Caption         =   "添加"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   480
      Width           =   1095
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
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton cmdQueryRooms 
         Caption         =   "确定"
         Height          =   360
         Left            =   3600
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtRoomName 
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   2535
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
   Begin MSComctlLib.ListView lvRooms 
      Height          =   6255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11033
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ckQuery_Click()
    EnableQuery CBool(Me.ckQuery.value)
    If Me.ckQuery.value = 0 Then
        PopulateRooms
    End If
End Sub

Private Sub cmdAddRoom_Click()
    If m_frmRoom Is Nothing Then Set m_frmRoom = New frmRoom
    m_frmRoom.m_action = MgrAction.add_
    m_frmRoom.Show 1
End Sub

Private Sub cmdDeleteRooms_Click()
    Dim lvItem As ListItem
    Dim del_ids As Variant
    For Each lvItem In Me.lvRooms.ListItems
        If lvItem.Checked Then
            AppendToVariantArr del_ids, Mid(lvItem.key, 4)
        End If
    Next
    If IsArray(del_ids) Then
        DeleteRoomsByIds del_ids
        RefreshRooms
    End If
End Sub

Private Sub cmdQueryRooms_Click()
    Dim data As Variant
    data = GetRoom()
    PopulateRooms data
End Sub

Private Sub cmdShowAllRooms_Click()
    PopulateRooms
End Sub

Private Sub Form_Load()
    With Me.lvRooms
        .ColumnHeaders.Add 1, "name", L_(LBL.room_name), 4000
    End With
    
    EnableQuery False
    
    PopulateRooms
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetRoom()
    Dim ret(Room.BOF_ + 1 To Room.EOF_ - 1) As Variant
    ret(Room.name_) = Trim$(Me.txtRoomName.Text)
    GetRoom = ret
End Function

Private Sub PopulateRooms(Optional ByVal Filter As Variant = Empty)
    Dim Rooms As Variant
    Dim i As Long
    Dim lvItem As ListItem
    Rooms = LoadRoomsFromDB(m_db, Filter)
    With Me.lvRooms
        .ListItems.Clear
        If Not IsEmpty(Rooms) Then
            For i = LBound(Rooms) To UBound(Rooms)
                Set lvItem = .ListItems.Add(i + 1, "id " & (Rooms(i)(Room.id)), Rooms(i)(Room.name_))
            Next
        End If
    End With
End Sub

Private Sub EnableQuery(ByVal Enabled As Boolean)
    Me.txtRoomName.Enabled = Enabled
    Me.cmdQueryRooms.Enabled = Enabled
End Sub

Public Sub RefreshRooms()
    If Me.ckQuery.value Then
        cmdQueryRooms_Click
    Else
        cmdShowAllRooms_Click
    End If
End Sub

Private Sub lvRooms_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lvRooms
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

