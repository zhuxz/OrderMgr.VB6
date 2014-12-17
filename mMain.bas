Attribute VB_Name = "mMain"
Option Explicit

Public m_MDIForm As MDIForm
Public m_frmConfig As frmConfig
Public m_frmEmployee As frmEmployee
Public m_frmRoom As frmRoom
Public m_frmOrder As frmOrder
Public m_frmOrders As frmOrders

Public m_ini As CIni
Public m_db As ADODB.Connection
Public m_dbSetting As mDefine.DatabaseSetting

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
    Dim source As String
    Dim ret As mDefine.DatabaseSetting
    
    source = Trim$(m_ini.ReadString(mDefine.INISEC_DBSETTING, mDefine.INIKEY_SOURCE, ""))
    If source = "" Then
        source = App.path & "\" & mDefine.DB_ACCESSFILENAME
    End If
    
    ret.source = source
    ret.pwd = Trim$(m_ini.ReadString(mDefine.INISEC_DBSETTING, mDefine.INIKEY_PWD, ""))
    ret.serverIP = Trim$(m_ini.ReadString(mDefine.INISEC_DBSETTING, mDefine.INIKEY_SERVERIP, ""))
    ret.account = Trim$(m_ini.ReadString(mDefine.INISEC_DBSETTING, mDefine.INIKEY_ACCOUNT, ""))
    ret.cnString = BuildDatabaseConnectString(source, ret.serverIP, ret.account, ret.pwd)
    
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
        .WriteString mDefine.INISEC_DBSETTING, mDefine.INIKEY_SOURCE, Setting.source
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

Public Sub ShowForm_Employee()
    If m_frmEmployee Is Nothing Then Set m_frmEmployee = New frmEmployee
    m_frmEmployee.Show: m_frmEmployee.SetFocus
    m_ini.WriteString mDefine.INISEC_MAIN, mDefine.INIKEY_ACTIVEFORM, mDefine.FORMNAME_EMPLOYEE
End Sub

Public Sub ShowForm_Room()
    If m_frmRoom Is Nothing Then Set m_frmRoom = New frmRoom
    m_frmRoom.Show: m_frmRoom.SetFocus
    m_ini.WriteString mDefine.INISEC_MAIN, mDefine.INIKEY_ACTIVEFORM, mDefine.FORMNAME_ROOM
End Sub

Public Sub ShowForm_Order()
    If m_frmOrder Is Nothing Then Set m_frmOrder = New frmOrder
    m_frmOrder.Show: m_frmOrder.SetFocus
    m_ini.WriteString mDefine.INISEC_MAIN, mDefine.INIKEY_ACTIVEFORM, mDefine.FORMNAME_ORDER
End Sub

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
