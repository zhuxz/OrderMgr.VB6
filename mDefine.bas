Attribute VB_Name = "mDefine"
Option Explicit

Public Const UI_MARGIN As Long = 20

Public Const MSG_TITLE As String = "小马过河账单管理系统"
Public Const MSG_INI_MISSING As String = "缺少配置文件!"
Public Const MSG_CONNECTDB_FAILED As String = "连接数据库失败."
Public Const MSG_CONNECTDB_SUCCESS As String = "连接数据库成功."
Public Const MSG_VALIDEMPLOYEENAME As String = "姓名不能为空."
Public Const MSG_ADDEMPLOYEEASK As String = "添加成功，是否继续添加？"
Public Const MSG_VALIDSERVICENAME As String = "名称不能为空."

Public Const DB_ACCESSFILENAME As String = "order.mdb"
Public Const DBTN_EMPLOYEES As String = "employees"
Public Const DBTN_SERVICES As String = "services"

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
Public Const FORMNAME_EMPLOYEES As String = "frmEmployees"
'Public Const FORMNAME_ROOM As String = "frmRoom"
Public Const FORMNAME_ORDER As String = "frmOrder"
Public Const FORMNAME_SERVICES As String = "frmServices"

Public Const SEX_MALE As String = "男"
Public Const SEX_FEMALE As String = "女"

Public Type DatabaseSetting
    serverIP As String
    Source As String
    account As String
    pwd As String
    cnString As String
End Type

Public Enum MgrAction
    add_ = 1
    update_
End Enum

Public Enum Employee
    BOF_
    ID
    name_
    sex
    EOF_
End Enum

Public Enum Service
    BOF_
    ID
    name_
    price
    EOF_
End Enum

Public Enum Room
    BOF_
    ID
    name_
    price
    EOF_
End Enum

Public Enum Order
    BOF_
    ID
    employeeId
    employeeName
    employeeSex
    roomId
    roomName
    serviceId
    serviceName
    price
    createDate
    memo_
    EOF_
End Enum

Public Enum LBL
    BOF_
    name_
    sex
    service_name
    room_name
    price
    createDate
    memo_
    EOF_
End Enum

Public Enum FRMN
    BOF_
    Rooms
    Orders
    EOF_
End Enum

Public Enum TBN
    BOF_
    Rooms
    Orders
    EOF_
End Enum

Public Enum MSG
    BOF_
    ValidRoomName
    ValidServiceName
    ValidServicePrice
    QueryDelete
    EOF_
End Enum
