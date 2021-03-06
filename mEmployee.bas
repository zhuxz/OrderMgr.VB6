Attribute VB_Name = "mEmployee"
Option Explicit

Public Function LoadEmployeesFromDB(Conn As ADODB.Connection, Optional ByVal Filter As Variant = Empty)
    Dim sql As String
    sql = "SELECT * FROM " & mDefine.DBTN_EMPLOYEES

    Dim arr() As String
    Dim n As Long

    If Not IsEmpty(Filter) Then
        If Len(Filter(Employee.name_)) > 0 Then
            n = n + 1
            ReDim Preserve arr(1 To n) As String
            arr(n) = "name like '%" & Filter(Employee.name_) & "%'"
        End If

        If Len(Filter(Employee.sex)) > 0 Then
            n = n + 1
            ReDim Preserve arr(1 To n) As String
            arr(n) = "sex='" & Filter(Employee.sex) & "'"
        End If

        If n > 0 Then
            sql = sql & " WHERE " & Join(arr, " AND ")
        End If
    End If

    Dim ret As Variant
    Dim item As Variant
    Dim rs As ADODB.Recordset

    Set rs = Conn.Execute(sql)
    If rs Is Nothing Then
        ''
    Else
        If rs.EOF Or rs.BOF Then
            ''
        Else
            Do
                item = VariantArr(Employee.BOF_ + 1, Employee.EOF_ - 1)
                item(Employee.id) = rs("ID").value
                item(Employee.name_) = rs("name").value
                item(Employee.sex) = rs("sex").value
                AppendToVariantArr ret, item
                rs.MoveNext
            Loop While Not rs.EOF
            LoadEmployeesFromDB = ret
        End If
    End If
End Function

Public Function DeleteEmployeesByIds(ByVal IDArr As Variant)
    If IsArray(IDArr) Then
        Dim ids As String
        Dim sql As String

        ids = Join(IDArr, ",")
        sql = "DELETE * FROM " & mDefine.DBTN_EMPLOYEES & " WHERE ID IN (" & ids & ")"
        m_db.Execute sql
    End If
End Function
