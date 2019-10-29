Imports System.Environment
Imports System.Data.SqlClient

Module MacolaStartup

    Friend cn As SqlConnection
    Public Sub MacStartup(ByVal db As String)
        'db = "DATA" ' Test Database

        Dim ConnStr As String = "Data Source=MASS_SQL2;Initial Catalog=" & db & ";Persist Security Info=True;User ID=sa;Password=STMARTIN"
        cn = New SqlConnection
        cn.ConnectionString = ConnStr
        cn.Open()

    End Sub


End Module
