Attribute VB_Name = "mOrder"
Option Explicit

Public Function LoadOrdersFromDB(Conn As ADODB.Connection, Optional ByVal Filter As Variant = Empty)
    Dim sql As String
    sql = "SELECT o.*," & _
        " e.name AS employeeName, e.sex as employeeSex," & _
        " r.name AS roomName," & _
        " s.desc as serviceName" & _
        " FROM ((" & TBN_(TBN.Orders) & " AS o" & _
        " LEFT OUTER JOIN " & mDefine.DBTN_EMPLOYEES & " AS e ON e.id=o.employeeId)" & _
        " LEFT OUTER JOIN " & mDefine.DBTN_SERVICES & " AS s ON s.id=o.serviceId)" & _
        " LEFT OUTER JOIN " & TBN_(TBN.Rooms) & " AS r ON r.id=o.roomId"

    Dim arr() As String
    Dim n As Long

    If Not IsEmpty(Filter) Then
        If Len(Filter(Order.employeeName)) > 0 Then
            n = n + 1
            ReDim Preserve arr(1 To n) As String
            arr(n) = "e.name LIKE '%" & Filter(Order.employeeName) & "%'"
        End If
        
        If Len(Filter(Order.employeeSex)) > 0 Then
            n = n + 1
            ReDim Preserve arr(1 To n) As String
            arr(n) = "e.sex ='" & Filter(Order.employeeSex) & "'"
        End If
        
        If Len(Filter(Order.serviceName)) > 0 Then
            n = n + 1
            ReDim Preserve arr(1 To n) As String
            arr(n) = "s.desc LIKE '%" & Filter(Order.serviceName) & "%'"
        End If
        
        If Len(Filter(Order.roomName)) > 0 Then
            n = n + 1
            ReDim Preserve arr(1 To n) As String
            arr(n) = "r.name LIKE '%" & Filter(Order.roomName) & "%'"
        End If
        
        If Len(Filter(Order.price)) > 0 Then
            n = n + 1
            ReDim Preserve arr(1 To n) As String
            arr(n) = "o.price =" & Filter(Order.price) & ""
        End If

        If n > 0 Then
            sql = sql & " WHERE " & Join(arr, " AND ")
        End If
    End If
    
    sql = sql & " ORDER BY o.ID DESC"

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
                item = VariantArr(Order.BOF_ + 1, Order.EOF_ - 1)
                item(Order.ID) = rs("ID").value
                If Not IsNull(rs("employeeId")) Then item(Order.employeeId) = rs("employeeId").value
                If Not IsNull(rs("employeeName")) Then item(Order.employeeName) = rs("employeeName").value
                If Not IsNull(rs("employeeSex")) Then item(Order.employeeSex) = rs("employeeSex").value
                If Not IsNull(rs("roomId")) Then item(Order.roomId) = rs("roomId").value
                If Not IsNull(rs("roomName")) Then item(Order.roomName) = rs("roomName").value
                If Not IsNull(rs("serviceId")) Then item(Order.serviceId) = rs("serviceId").value
                If Not IsNull(rs("serviceName")) Then item(Order.serviceName) = rs("serviceName").value
                If Not IsNull(rs("price")) Then item(Order.price) = rs("price").value
                If Not IsNull(rs("createDate")) Then item(Order.createDate) = rs("createDate").value
                If Not IsNull(rs("memo")) Then item(Order.memo_) = rs("memo").value
                AppendToVariantArr ret, item
                rs.MoveNext
            Loop While Not rs.EOF
            LoadOrdersFromDB = ret
        End If
    End If
End Function

Public Function DeleteOrdersByIds(ByVal IDArr As Variant)
    If IsArray(IDArr) Then
        Dim ids As String
        Dim sql As String

        ids = Join(IDArr, ",")
        sql = "DELETE * FROM " & TBN_(TBN.Orders) & " WHERE ID IN (" & ids & ")"
        m_db.Execute sql
    End If
End Function

