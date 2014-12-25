VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmServices 
   Caption         =   "服务项目"
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
   ScaleHeight     =   19087.8
   ScaleMode       =   0  'User
   ScaleWidth      =   11925
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAddService 
      Cancel          =   -1  'True
      Caption         =   "添加"
      Height          =   360
      Left            =   5160
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin VB.CheckBox ckQuery 
      Caption         =   "查询"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   855
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
      TabIndex        =   3
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtServicePrice 
         Height          =   360
         Left            =   840
         TabIndex        =   9
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtServiceName 
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton cmdQueryServices 
         Caption         =   "确定"
         Height          =   360
         Left            =   3600
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "项目："
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "单价："
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdDeleteServices 
      Caption         =   "删除"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdShowAllServices 
      Caption         =   "显示全部"
      Height          =   375
      Left            =   10560
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvServices 
      Height          =   5895
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10398
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
Attribute VB_Name = "frmServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ckQuery_Click()
    EnableQuery CBool(Me.ckQuery.value)
    If Me.ckQuery.value = 0 Then
        PopulateServices
    End If
End Sub

Private Sub cmdAddService_Click()
    If m_frmService Is Nothing Then Set m_frmService = New frmService
    m_frmService.m_action = MgrAction.add_
    m_frmService.Show 1
End Sub

Private Sub cmdDeleteServices_Click()
    Dim lvItem As ListItem
    Dim del_ids As Variant
    For Each lvItem In Me.lvServices.ListItems
        If lvItem.Checked Then
            AppendToVariantArr del_ids, Mid(lvItem.key, 4)
        End If
    Next
    If IsArray(del_ids) Then
        DeleteServicesByIds del_ids
        RefreshServices
    End If
End Sub

Private Sub cmdQueryServices_Click()
    Dim data As Variant
    data = GetService()
    PopulateServices data
End Sub

Private Sub cmdShowAllServices_Click()
    PopulateServices
End Sub

Private Sub Form_Load()
    With Me.lvServices
        .ColumnHeaders.Add 1, "desc", L_(LBL.service_name), 4000
        .ColumnHeaders.Add 2, "price", L_(LBL.price), 1500
    End With
    
    EnableQuery False
    
    PopulateServices
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetService()
    Dim ret(Service.BOF_ + 1 To Service.EOF_ - 1) As Variant
    ret(Service.name_) = Trim$(Me.txtServiceName.Text)
    ret(Service.price) = Me.txtServicePrice.Text
    GetService = ret
End Function

Private Sub PopulateServices(Optional ByVal Filter As Variant = Empty)
    Dim services As Variant
    Dim i As Long
    Dim lvItem As ListItem
    services = LoadServicesFromDB(m_db, Filter)
    With Me.lvServices
        .ListItems.Clear
        If Not IsEmpty(services) Then
            For i = LBound(services) To UBound(services)
                Set lvItem = .ListItems.Add(i + 1, "id " & (services(i)(Service.id)), services(i)(Service.name_))
                lvItem.SubItems(1) = services(i)(Service.price)
            Next
        End If
    End With
End Sub

Private Sub EnableQuery(ByVal Enabled As Boolean)
    Me.txtServiceName.Enabled = Enabled
    Me.txtServicePrice.Enabled = Enabled
    Me.cmdQueryServices.Enabled = Enabled
End Sub

Public Sub RefreshServices()
    If Me.ckQuery.value Then
        cmdQueryServices_Click
    Else
        cmdShowAllServices_Click
    End If
End Sub

Private Sub lvServices_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lvServices
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
