Attribute VB_Name = "mMain"
Option Explicit

Public m_MDIForm As MDIForm
Public m_frmConfig As frmConfig
Public m_frmEmployees As frmEmployees
Public m_frmOrder As frmOrder
Public m_frmOrders As frmOrders
Public m_frmServices As frmServices
Public m_frmRooms As frmRooms

Public m_frmEmployee As frmEmployee
Public m_frmService As frmService
Public m_frmRoom As frmRoom

Public m_ini As CIni
Public m_db As ADODB.Connection
Public m_dbSetting As mDefine.DatabaseSetting

Public L_ As Variant
Public FN_ As Variant
Public TBN_ As Variant
Public MSG_ As Variant

Sub Main()
    Set m_ini = New CIni
    If Not m_ini.checkIni() Then
        MsgBox mDefine.MSG_INI_MISSING, vbExclamation, mDefine.MSG_TITLE
    End If
    
    m_dbSetting = GetDatabaseSetting()
    Set m_db = ConnectDatabase(m_dbSetting.cnString)
    If m_db Is Nothing Then
        MsgBox mDefine.MSG_CONNECTDB_FAILED, vbExclamation, mDefine.MSG_TITLE
    End If
    
    L_ = mMain.BuildLabelCollection()
    FN_ = mMain.BuildFormNameCollection()
    TBN_ = mMain.BuildTableNameCollection()
    MSG_ = mMain.BuildMessageCollection()
    
    Set m_MDIForm = New MDIMain
    m_MDIForm.Show
End Sub

Public Function ConnectDatabase(ByVal ConnectString As String) As ADODB.Connection
On Error GoTo eh
    Dim cn As New ADODB.Connection
    cn.ConnectionTimeout = m_ini.ReadNumber(mDefine.INISEC_DBSETTING, mDefine.INIKEY_TIMEOUT, 60)
    cn.Open ConnectString
    Set ConnectDatabase = cn
eh:
    Err.Clear
End Function

Public Function GetDatabaseSetting() As DatabaseSetting
    Dim Source As String
    Dim ret As mDefine.DatabaseSetting
    
    Source = Trim$(m_ini.ReadString(mDefine.INISEC_DBSETTING, mDefine.INIKEY_SOURCE, ""))
    If Source = "" Then
        Source = App.path & "\" & mDefine.DB_ACCESSFILENAME
    End If
    
    ret.Source = Source
    ret.pwd = Trim$(m_ini.ReadString(mDefine.INISEC_DBSETTING, mDefine.INIKEY_PWD, ""))
    ret.serverIP = Trim$(m_ini.ReadString(mDefine.INISEC_DBSETTING, mDefine.INIKEY_SERVERIP, ""))
    ret.account = Trim$(m_ini.ReadString(mDefine.INISEC_DBSETTING, mDefine.INIKEY_ACCOUNT, ""))
    ret.cnString = BuildDatabaseConnectString(Source, ret.serverIP, ret.account, ret.pwd)
    
    GetDatabaseSetting = ret
End Function

Public Function BuildDatabaseConnectString(Optional ByVal DBSource As String, _
    Optional ByVal DBSeverIP As String, _
    Optional ByVal DBAccount As String, _
    Optional ByVal DBPwd As String)
    Dim ret As String
    ret = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBSource '& ";" & "Database password=" & DBPwd & ""
    'ret = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & DBSeverIP & _
        ";Initial Catalog=" & DBSource & _
        ";USER ID=" & DBAccount & _
        ";PASSWORD=" & DBPwd & ";"
    BuildDatabaseConnectString = ret
End Function

Public Sub SaveDatabaseSetting(ByRef Setting As mDefine.DatabaseSetting)
    With m_ini
        .WriteString mDefine.INISEC_DBSETTING, mDefine.INIKEY_SERVERIP, Setting.serverIP
        .WriteString mDefine.INISEC_DBSETTING, mDefine.INIKEY_SOURCE, Setting.Source
        .WriteString mDefine.INISEC_DBSETTING, mDefine.INIKEY_ACCOUNT, Setting.account
        .WriteString mDefine.INISEC_DBSETTING, mDefine.INIKEY_PWD, Setting.pwd
    End With
End Sub

Public Sub ShowForm_Config()
    If m_frmConfig Is Nothing Then Set m_frmConfig = New frmConfig
    m_frmConfig.Show: m_frmConfig.SetFocus
    m_ini.WriteString mDefine.INISEC_MAIN, mDefine.INIKEY_ACTIVEFORM, mDefine.FORMNAME_CONFIG
End Sub

Public Sub ShowForm_Orders()
    If m_frmOrders Is Nothing Then Set m_frmOrders = New frmOrders
    m_frmOrders.Show: m_frmOrders.SetFocus
    m_ini.WriteString mDefine.INISEC_MAIN, mDefine.INIKEY_ACTIVEFORM, mDefine.FORMNAME_ORDERS
End Sub

Public Sub ShowForm_Employees()
    If m_frmEmployees Is Nothing Then Set m_frmEmployees = New frmEmployees
    m_frmEmployees.Show: m_frmEmployees.SetFocus
    m_ini.WriteString mDefine.INISEC_MAIN, mDefine.INIKEY_ACTIVEFORM, mDefine.FORMNAME_EMPLOYEES
End Sub

Public Sub ShowForm_Rooms()
    If m_frmRooms Is Nothing Then Set m_frmRooms = New frmRooms
    m_frmRooms.Show: m_frmRooms.SetFocus
    m_ini.WriteString mDefine.INISEC_MAIN, mDefine.INIKEY_ACTIVEFORM, FN_(FRMN.Rooms)
End Sub

Public Sub ShowForm_Order()
    If m_frmOrder Is Nothing Then Set m_frmOrder = New frmOrder
    m_frmOrder.Show: m_frmOrder.SetFocus
    m_ini.WriteString mDefine.INISEC_MAIN, mDefine.INIKEY_ACTIVEFORM, mDefine.FORMNAME_ORDER
End Sub

Public Sub ShowForm_Services()
    If m_frmServices Is Nothing Then Set m_frmServices = New frmServices
    m_frmServices.Show: m_frmServices.SetFocus
    m_ini.WriteString mDefine.INISEC_MAIN, mDefine.INIKEY_ACTIVEFORM, mDefine.FORMNAME_SERVICES
End Sub

Public Function BuildLabelCollection()
    Dim ret(LBL.BOF_ + 1 To LBL.EOF_ - 1) As String
    ret(LBL.name_) = "姓名"
    ret(LBL.sex) = "性别"
    ret(LBL.service_name) = "项目名称"
    ret(LBL.price) = "单价（元）"
    ret(LBL.room_name) = "房号"
    ret(LBL.createDate) = "创建日期"
    ret(LBL.memo_) = "备注"
    BuildLabelCollection = ret
End Function

Public Function BuildFormNameCollection()
    Dim ret(FRMN.BOF_ + 1 To FRMN.EOF_ - 1) As String
    ret(FRMN.Rooms) = "frmRooms"
    ret(FRMN.Orders) = "frmRooms"
    BuildFormNameCollection = ret
End Function

Public Function BuildTableNameCollection()
    Dim ret(TBN.BOF_ + 1 To TBN.EOF_ - 1) As String
    ret(TBN.Rooms) = "rooms"
    ret(TBN.Orders) = "orders"
    BuildTableNameCollection = ret
End Function

Public Function BuildMessageCollection()
    Dim ret(MSG.BOF_ + 1 To MSG.EOF_ - 1) As String
    ret(MSG.ValidRoomName) = "房号不能空."
    ret(MSG.ValidServiceName) = "服务名称不能空."
    ret(MSG.ValidServicePrice) = "单价不能空."
    BuildMessageCollection = ret
End Function

'Public Function IsFormExists(ByVal FormName As String) As Boolean
'    Dim ctl As Object
'    Dim ret As Boolean
'    For Each ctl In MDIMain.Controls
'        If ctl.Name = FormName Then
'            ret = True
'            Exit For
'        End If
'    Next
'    IsFormExists = ret
'End Function
