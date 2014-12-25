Attribute VB_Name = "mService"
Option Explicit

Public Function LoadServicesFromDB(Conn As ADODB.Connection, Optional ByVal Filter As Variant = Empty)
    Dim sql As String
    sql = "SELECT * FROM " & mDefine.DBTN_SERVICES

    Dim arr() As String
    Dim n As Long

    If Not IsEmpty(Filter) Then
        If Len(Filter(Service.name_)) > 0 Then
            n = n + 1
            ReDim Preserve arr(1 To n) As String
            arr(n) = "desc like '%" & Filter(Service.name_) & "%'"
        End If

        If Len(Filter(Service.price)) > 0 Then
            n = n + 1
            ReDim Preserve arr(1 To n) As String
            arr(n) = "price=" & Filter(Service.price)
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
                item = VariantArr(Service.BOF_ + 1, Service.EOF_ - 1)
                item(Service.id) = rs("ID").value
                item(Service.name_) = rs("desc").value
                item(Service.price) = rs("price").value
                AppendToVariantArr ret, item
                rs.MoveNext
            Loop While Not rs.EOF
            LoadServicesFromDB = ret
        End If
    End If
End Function

Public Function DeleteServicesByIds(ByVal IDArr As Variant)
    If IsArray(IDArr) Then
        Dim ids As String
        Dim sql As String

        ids = Join(IDArr, ",")
        sql = "DELETE * FROM " & mDefine.DBTN_SERVICES & " WHERE ID IN (" & ids & ")"
        m_db.Execute sql
    End If
End Function
