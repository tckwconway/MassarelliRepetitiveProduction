Imports System.Data.SqlClient
''' <summary>
''' DAC: Data Access
''' </summary>
''' <remarks></remarks>


Public Class DAC

    Public Shared Function ExecuteSP_RetVal(ByVal storedprocedure As String, _
    ByVal cn As SqlConnection, ByVal ParamArray arrParam() As SqlParameter) As Integer
        Dim cmd As New SqlCommand
        cmd.Connection = cn
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = storedprocedure
        cmd.Parameters.Clear()
        Dim i As Integer = 0
        If arrParam IsNot Nothing Then
            For Each param As SqlParameter In arrParam
                Debug.Print(CStr(arrParam(i).Value))
                cmd.Parameters.Add(param)
                i = i + 1
            Next
        End If

        cmd.ExecuteNonQuery()
        Dim ReturnValue As String
        Dim retval As String = "oNewKey"
        ReturnValue = cmd.Parameters(retval).Value
        Return ReturnValue


    End Function
    Public Shared Function ExecuteSP_ReturnObject(ByVal storedprocedure As String, _
        ByVal cn As SqlConnection, ByVal ParamArray arrParam() As SqlParameter) As Integer
        Dim cmd As New SqlCommand
        cmd.Connection = cn
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = storedprocedure
        cmd.Parameters.Clear()
        Dim i As Integer = 0
        If arrParam IsNot Nothing Then
            For Each param As SqlParameter In arrParam
                Debug.Print(CStr(arrParam(i).Value))
                cmd.Parameters.Add(param)
                i = i + 1
            Next
        End If
        Dim retval As Object = cmd.ExecuteNonQuery()
        Return retval

        'cmd.ExecuteNonQuery()
        'Dim ReturnValue As String
        'Dim retval As Object
        'ReturnValue = cmd.Parameters(retval).Value
        'Return ReturnValue


    End Function

    Public Shared Function ExecuteSP(ByVal storedprocedure As String, ByVal cn As SqlConnection, ByVal ParamArray arrParam() As SqlParameter) As Integer
        Dim cmd As New SqlCommand
        cmd.Connection = cn
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = storedprocedure
        cmd.Parameters.Clear()
        Dim i As Integer = 0
        If arrParam IsNot Nothing Then
            For Each param As SqlParameter In arrParam
                Debug.Print(CStr(arrParam(i).Value))
                cmd.Parameters.Add(param)
                i = i + 1
            Next
        End If
        Try
            cmd.ExecuteNonQuery()

        Catch ex As Exception

        End Try
        Dim ReturnValue As Integer = 0
        ''Dim ReturnValue As String
        ''Dim retval As String = "oNewKey"
        ''ReturnValue = cmd.Parameters(retval).Value
        Return ReturnValue


    End Function
    Public Shared Function ExecuteSP_DataTable(ByVal storedprocedure As String, _
        ByVal cn As SqlConnection, ByVal ParamArray arrParam() As SqlParameter) As DataTable
        Dim dt As DataTable
        Dim cmd As New SqlCommand
        cmd.Connection = cn
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = storedprocedure
        cmd.Parameters.Clear()
        Dim i As Integer = 0
        If arrParam IsNot Nothing Then
            For Each param As SqlParameter In arrParam
                Debug.Print(CStr(arrParam(i).Value))
                cmd.Parameters.Add(param)
                i = i + 1
            Next
        End If
        Dim da As New SqlDataAdapter(cmd)
        dt = New DataTable
        da.Fill(dt)

        Return dt

    End Function

    Public Shared Function ExecuteSaveSP(ByVal storedprocedure As String, _
       ByVal cn As SqlConnection, ByVal ParamArray arrParam() As SqlParameter) As Integer
        Dim cmd As New SqlCommand
        cmd.Connection = cn
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = storedprocedure
        cmd.Parameters.Clear()
        Dim i As Integer = 0
        If arrParam IsNot Nothing Then
            For Each param As SqlParameter In arrParam
                'Debug.Print(CStr(arrParam(i).Value))
                cmd.Parameters.Add(param)
                i = i + 1
            Next
        End If
        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception

        End Try
        Dim ReturnValue As Integer = 0
        'Boolean = True
        'Dim retval As String = "oNewKey"
        'ReturnValue = cmd.Parameters(retval).Value
        'Return ReturnValue
        For Each param As SqlParameter In arrParam
            If param.Direction = ParameterDirection.InputOutput Or param.Direction = ParameterDirection.Output Then
                Return param.Value
            End If
        Next

        Return ReturnValue


    End Function
    Public Shared Function ExecuteSP_Reader(ByVal storedprocedure As String, _
    ByVal cn As SqlConnection, ByVal ParamArray arrParam() As SqlParameter) As SqlDataReader
        Dim cmd As New SqlCommand
        cmd.Connection = cn
        cmd.CommandTimeout = 5000
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = storedprocedure
        cmd.Parameters.Clear()
        Dim i As Integer = 0
        If arrParam IsNot Nothing Then
            For Each param As SqlParameter In arrParam
                Debug.Print(CStr(arrParam(i).Value))
                cmd.Parameters.Add(param)
                i = i + 1
            Next
        End If
        Dim rd As SqlDataReader
        'Dim ReturnValue As String
        'Dim retval As String = "oNewKey"
        'ReturnValue = cmd.Parameters(retval).Value
        'Return ReturnValue
        Try
            rd = cmd.ExecuteReader
            Return rd
        Catch ex As Exception

            MsgBox(ex.Message, MsgBoxStyle.OkOnly, " Error ")
            Return Nothing
        End Try


    End Function
    Public Shared Sub Execute_NonSQL(ByVal sSQL As String, ByVal cn As SqlConnection)

        Dim cmd As New SqlCommand
        cmd.Connection = cn
        cmd.CommandType = CommandType.Text
        cmd.CommandText = sSQL
        cmd.ExecuteNonQuery()

    End Sub

    Public Shared Function Execute_SQL(ByVal sSQL As String, ByVal cn As SqlConnection) As DataTable

        Return Nothing

    End Function
    Public Shared Function Execute_Scalar(ByVal sSQL As String, ByVal cn As SqlConnection) As Object

        Dim cmd As New SqlCommand
        cmd.Connection = cn
        cmd.CommandType = CommandType.Text
        cmd.CommandText = sSQL
        Dim retval As Object
        retval = cmd.ExecuteScalar()
        Return retval

    End Function
    Public Shared Function ExecuteSQL_DataTable(ByVal sSQL As String, ByVal cn As SqlConnection, ByVal tableName As String) As DataTable
        Dim dt As DataTable
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter(sSQL, cn)
        da.Fill(ds)
        dt = ds.Tables(0)
        dt.TableName = tableName
        Return dt

    End Function


    Public Shared Function ExecuteSQL_Reader(ByVal sSQL As String, ByVal cn As SqlConnection) As SqlDataReader
        Dim reader As SqlDataReader

        'If reader.IsClosed = False Then reader.Close()
        Dim cmd As New SqlCommand
        cmd.CommandType = CommandType.Text
        cmd.CommandText = sSQL
        cmd.Connection = cn
        reader = cmd.ExecuteReader()

        Return reader

    End Function

    Public Shared Function Parameter(ByVal parameterName As String, ByVal parameterValue As Object, _
                                     ByVal parameterDirection As Integer) As SqlParameter

        Dim param As New SqlParameter
        param.ParameterName = parameterName
        param.Value = parameterValue
        param.Direction = parameterDirection
        Return param

    End Function

    Public Shared Function Parameter(ByVal parameterName As String, ByVal parameterValue As Object, _
                                     ByVal parameterDirection As Integer, ByVal parameterSQLDbType As SqlDbType) As Object

        Dim param As New SqlParameter
        param.ParameterName = parameterName
        param.Value = parameterValue
        param.Direction = parameterDirection
        param.SqlDbType = parameterSQLDbType
        Return param

    End Function
End Class
