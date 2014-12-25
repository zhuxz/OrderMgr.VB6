Attribute VB_Name = "mRoom"
Option Explicit

Public Function LoadRoomsFromDB(Conn As ADODB.Connection, Optional ByVal Filter As Variant = Empty)
    Dim sql As String
    sql = "SELECT * FROM " & TBN_(TBN.Rooms)

    Dim arr() As String
    Dim n As Long

    If Not IsEmpty(Filter) Then
        If Len(Filter(Room.name_)) > 0 Then
            n = n + 1
            ReDim Preserve arr(1 To n) As String
            arr(n) = "name like '%" & Filter(Room.name_) & "%'"
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
                item = VariantArr(Room.BOF_ + 1, Room.EOF_ - 1)
                item(Room.id) = rs("ID").value
                item(Room.name_) = rs("name").value
                AppendToVariantArr ret, item
                rs.MoveNext
            Loop While Not rs.EOF
            LoadRoomsFromDB = ret
        End If
    End If
End Function

Public Function DeleteRoomsByIds(ByVal IDArr As Variant)
    If IsArray(IDArr) Then
        Dim ids As String
        Dim sql As String

        ids = Join(IDArr, ",")
        sql = "DELETE * FROM " & TBN_(TBN.Rooms) & " WHERE ID IN (" & ids & ")"
        m_db.Execute sql
    End If
End Function
