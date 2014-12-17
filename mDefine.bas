Attribute VB_Name = "mDefine"
Option Explicit

Public Const UI_MARGIN As Long = 20

Public Const MSG_TITLE As String = "С������˵�����ϵͳ"
Public Const MSG_INI_MISSING As String = "ȱ�������ļ�!"
Public Const MSG_CONNECTDB_FAILED As String = "�������ݿ�ʧ��."
Public Const MSG_CONNECTDB_SUCCESS As String = "�������ݿ�ɹ�."
Public Const MSG_NAMEVALID As String = "��������Ϊ��."

Public Const DB_ACCESSFILENAME As String = "order.mdb"
Public Const DBTN_EMPLOYEES As String = "employees"


Public Const INISEC_DBSETTING As String = "dbsetting"
Public Const INIKEY_TIMEOUT As String = "timeout"
Public Const INIKEY_SOURCE As String = "source"
Public Const INIKEY_SERVERIP As String = "serverIP"
Public Const INIKEY_ACCOUNT As String = "account"
Public Const INIKEY_PWD As String = "pwd"

Public Const INI_FILENAME As String = "app.ini"
Public Const INISEC_MAIN As String = "main"
Public Const INIKEY_ACTIVEFORM As String = "activeform"

Public Const FORMNAME_CONFIG As String = "frmConfig"
Public Const FORMNAME_ORDERS As String = "frmOrders"
Public Const FORMNAME_EMPLOYEE As String = "frmEmployee"
Public Const FORMNAME_ROOM As String = "frmRoom"
Public Const FORMNAME_ORDER As String = "frmOrder"

Public Const SEX_MALE As String = "��"
Public Const SEX_FEMALE As String = "Ů"
Public Const SEX_MALEID As Long = 1
Public Const SEX_FEMALEID As Long = 0

Public Type DatabaseSetting
    serverIP As String
    source As String
    account As String
    pwd As String
    cnString As String
End Type

Public Enum Employee
    BOF_
    name_
    sex
    EOF_
End Enum
