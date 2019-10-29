Imports System.ComponentModel
Imports System.Data.DataRowView
Imports System.Data.SqlClient
Imports System.Data
Imports System.Text
Imports System.Xml
Imports System.IO
Imports System.Data.SqlTypes

Public Class cSystemPeriodSettings

    Public prd_key As String
    Public str_dt_1 As Integer
    Public end_dt_1 As Integer
    Public str_dt_2 As Integer
    Public end_dt_2 As Integer
    Public str_dt_3 As Integer
    Public end_dt_3 As Integer
    Public str_dt_4 As Integer
    Public end_dt_4 As Integer
    Public str_dt_5 As Integer
    Public end_dt_5 As Integer
    Public str_dt_6 As Integer
    Public end_dt_6 As Integer
    Public str_dt_7 As Integer
    Public end_dt_7 As Integer
    Public str_dt_8 As Integer
    Public end_dt_8 As Integer
    Public str_dt_9 As Integer
    Public end_dt_9 As Integer
    Public str_dt_10 As Integer
    Public end_dt_10 As Integer
    Public str_dt_11 As Integer
    Public end_dt_11 As Integer
    Public str_dt_12 As Integer
    Public end_dt_12 As Integer
    Public str_dt_13 As Integer
    Public end_dt_13 As Integer
    Public no_of_val_prd As Integer
    Public def_prd_cd As String
    Public tmp_yr_cls_fg As String
    Public warn_or_prev As String
    Public current_prd As Integer
    Public wrnprv_glprd_fg_1 As String
    Public wrnprv_glprd_fg_2 As String
    Public wrnprv_glprd_fg_3 As String
    Public wrnprv_glprd_fg_4 As String
    Public wrnprv_glprd_fg_5 As String
    Public wrnprv_glprd_fg_6 As String
    Public wrnprv_glprd_fg_7 As String
    Public wrnprv_glprd_fg_8 As String
    Public wrnprv_glprd_fg_9 As String
    Public wrnprv_glprdfg_10 As String
    Public wrnprv_glprdfg_11 As String
    Public wrnprv_glprdfg_12 As String
    Public wrnprv_glprdfg_13 As String
    Public filler_0001 As String
    Public current_year As Integer

    Const TABLE As String = "SYPRDFIL_SQL"

    Public Enum PeriodSetting As Integer
        Inventory = 8
        Unknown = 0
    End Enum
    Private Function NewSystemPeriodSetting(ByVal m_prd_key As String, ByVal m_str_dt_1 As Integer, ByVal m_end_dt_1 As Integer, ByVal m_str_dt_2 As Integer, ByVal m_end_dt_2 As Integer, _
                                            ByVal m_str_dt_3 As Integer, ByVal m_end_dt_3 As Integer, ByVal m_str_dt_4 As Integer, ByVal m_end_dt_4 As Integer, ByVal m_str_dt_5 As Integer, _
                                            ByVal m_end_dt_5 As Integer, ByVal m_str_dt_6 As Integer, ByVal m_end_dt_6 As Integer, ByVal m_str_dt_7 As Integer, ByVal m_end_dt_7 As Integer, _
                                            ByVal m_str_dt_8 As Integer, ByVal m_end_dt_8 As Integer, ByVal m_str_dt_9 As Integer, ByVal m_end_dt_9 As Integer, ByVal m_str_dt_10 As Integer, _
                                            ByVal m_end_dt_10 As Integer, ByVal m_str_dt_11 As Integer, ByVal m_end_dt_11 As Integer, ByVal m_str_dt_12 As Integer, ByVal m_end_dt_12 As Integer, _
                                            ByVal m_str_dt_13 As Integer, ByVal m_end_dt_13 As Integer, ByVal m_no_of_val_prd As Integer, ByVal m_def_prd_cd As String, ByVal m_tmp_yr_cls_fg As String, _
                                            ByVal m_warn_or_prev As String, ByVal m_current_prd As Integer, ByVal m_wrnprv_glprd_fg_1 As String, ByVal m_wrnprv_glprd_fg_2 As String, ByVal m_wrnprv_glprd_fg_3 As String, _
                                            ByVal m_wrnprv_glprd_fg_4 As String, ByVal m_wrnprv_glprd_fg_5 As String, ByVal m_wrnprv_glprd_fg_6 As String, ByVal m_wrnprv_glprd_fg_7 As String, ByVal m_wrnprv_glprd_fg_8 As String, _
                                            ByVal m_wrnprv_glprd_fg_9 As String, ByVal m_wrnprv_glprdfg_10 As String, ByVal m_wrnprv_glprdfg_11 As String, ByVal m_wrnprv_glprdfg_12 As String, _
                                            ByVal m_wrnprv_glprdfg_13 As String, ByVal m_filler_0001 As String, m_current_year As Integer) As cSystemPeriodSettings
        Dim syprd As New cSystemPeriodSettings
        syprd.prd_key = m_prd_key
        syprd.str_dt_1 = m_str_dt_1
        syprd.end_dt_1 = m_end_dt_1
        syprd.str_dt_2 = m_str_dt_2
        syprd.end_dt_2 = m_end_dt_2
        syprd.str_dt_3 = m_str_dt_3
        syprd.end_dt_3 = m_end_dt_3
        syprd.str_dt_4 = m_str_dt_4
        syprd.end_dt_4 = m_end_dt_4
        syprd.str_dt_5 = m_str_dt_5
        syprd.end_dt_5 = m_end_dt_5
        syprd.str_dt_6 = m_str_dt_6
        syprd.end_dt_6 = m_end_dt_6
        syprd.str_dt_7 = m_str_dt_7
        syprd.end_dt_7 = m_end_dt_7
        syprd.str_dt_8 = m_str_dt_8
        syprd.end_dt_8 = m_end_dt_8
        syprd.str_dt_9 = m_str_dt_9
        syprd.end_dt_9 = m_end_dt_9
        syprd.str_dt_10 = m_str_dt_10
        syprd.end_dt_10 = m_end_dt_10
        syprd.str_dt_11 = m_str_dt_11
        syprd.end_dt_11 = m_end_dt_11
        syprd.str_dt_12 = m_str_dt_12
        syprd.end_dt_12 = m_end_dt_12
        syprd.str_dt_13 = m_str_dt_13
        syprd.end_dt_13 = m_end_dt_13
        syprd.no_of_val_prd = m_no_of_val_prd
        syprd.def_prd_cd = m_def_prd_cd
        syprd.tmp_yr_cls_fg = m_tmp_yr_cls_fg
        syprd.warn_or_prev = m_warn_or_prev
        syprd.current_prd = m_current_prd
        syprd.wrnprv_glprd_fg_1 = m_wrnprv_glprd_fg_1
        syprd.wrnprv_glprd_fg_2 = m_wrnprv_glprd_fg_2
        syprd.wrnprv_glprd_fg_3 = m_wrnprv_glprd_fg_3
        syprd.wrnprv_glprd_fg_4 = m_wrnprv_glprd_fg_4
        syprd.wrnprv_glprd_fg_5 = m_wrnprv_glprd_fg_5
        syprd.wrnprv_glprd_fg_6 = m_wrnprv_glprd_fg_6
        syprd.wrnprv_glprd_fg_7 = m_wrnprv_glprd_fg_7
        syprd.wrnprv_glprd_fg_8 = m_wrnprv_glprd_fg_8
        syprd.wrnprv_glprd_fg_9 = m_wrnprv_glprd_fg_9
        syprd.wrnprv_glprdfg_10 = m_wrnprv_glprdfg_10
        syprd.wrnprv_glprdfg_11 = m_wrnprv_glprdfg_11
        syprd.wrnprv_glprdfg_12 = m_wrnprv_glprdfg_12
        syprd.wrnprv_glprdfg_13 = m_wrnprv_glprdfg_13
        syprd.filler_0001 = m_filler_0001
        syprd.current_year = m_current_year
        Return syprd

    End Function


    Public Function GetSystemPeriodSetting(ByVal PeriodSetting_Key As Integer, ByVal cn As SqlConnection) As cSystemPeriodSettings

        'PRD_KEY_1 for Massarelli INVENTORY is "8"

        Dim syprd As New cSystemPeriodSettings
        Dim dt As New DataTable
        Dim sSQL As String = "Select * from " & TABLE & " " _
                           & "Where prd_key = '" & PeriodSetting_Key & "' "

        dt = DAC.ExecuteSQL_DataTable(sSQL, cn, TABLE)

        syprd = PopulateSystemPeriodSettings(dt)

        Return syprd

    End Function

    Private Function PopulateSystemPeriodSettings(ByVal dt As DataTable) As cSystemPeriodSettings
        Dim syprd As New cSystemPeriodSettings
        Dim obj(46) As Object
        For Each rw As DataRow In dt.Rows

            obj(0) = ReplaceNullValue(rw("prd_key"), dt.Columns("prd_key").DataType.ToString)
            obj(1) = ReplaceNullValue(rw("str_dt_1"), dt.Columns("str_dt_1").DataType.ToString)
            obj(2) = ReplaceNullValue(rw("end_dt_1"), dt.Columns("end_dt_1").DataType.ToString)
            obj(3) = ReplaceNullValue(rw("str_dt_2"), dt.Columns("str_dt_2").DataType.ToString)
            obj(4) = ReplaceNullValue(rw("end_dt_2"), dt.Columns("end_dt_2").DataType.ToString)
            obj(5) = ReplaceNullValue(rw("str_dt_3"), dt.Columns("str_dt_3").DataType.ToString)
            obj(6) = ReplaceNullValue(rw("end_dt_3"), dt.Columns("end_dt_3").DataType.ToString)
            obj(7) = ReplaceNullValue(rw("str_dt_4"), dt.Columns("str_dt_4").DataType.ToString)
            obj(8) = ReplaceNullValue(rw("end_dt_4"), dt.Columns("end_dt_4").DataType.ToString)
            obj(9) = ReplaceNullValue(rw("str_dt_5"), dt.Columns("str_dt_5").DataType.ToString)
            obj(10) = ReplaceNullValue(rw("end_dt_5"), dt.Columns("end_dt_5").DataType.ToString)
            obj(11) = ReplaceNullValue(rw("str_dt_6"), dt.Columns("str_dt_6").DataType.ToString)
            obj(12) = ReplaceNullValue(rw("end_dt_6"), dt.Columns("end_dt_6").DataType.ToString)
            obj(13) = ReplaceNullValue(rw("str_dt_7"), dt.Columns("str_dt_7").DataType.ToString)
            obj(14) = ReplaceNullValue(rw("end_dt_7"), dt.Columns("end_dt_7").DataType.ToString)
            obj(15) = ReplaceNullValue(rw("str_dt_8"), dt.Columns("str_dt_8").DataType.ToString)
            obj(16) = ReplaceNullValue(rw("end_dt_8"), dt.Columns("end_dt_8").DataType.ToString)
            obj(17) = ReplaceNullValue(rw("str_dt_9"), dt.Columns("str_dt_9").DataType.ToString)
            obj(18) = ReplaceNullValue(rw("end_dt_9"), dt.Columns("end_dt_9").DataType.ToString)
            obj(19) = ReplaceNullValue(rw("str_dt_10"), dt.Columns("str_dt_10").DataType.ToString)
            obj(20) = ReplaceNullValue(rw("end_dt_10"), dt.Columns("end_dt_10").DataType.ToString)
            obj(21) = ReplaceNullValue(rw("str_dt_11"), dt.Columns("str_dt_11").DataType.ToString)
            obj(22) = ReplaceNullValue(rw("end_dt_11"), dt.Columns("end_dt_11").DataType.ToString)
            obj(23) = ReplaceNullValue(rw("str_dt_12"), dt.Columns("str_dt_12").DataType.ToString)
            obj(24) = ReplaceNullValue(rw("end_dt_12"), dt.Columns("end_dt_12").DataType.ToString)
            obj(25) = ReplaceNullValue(rw("str_dt_13"), dt.Columns("str_dt_13").DataType.ToString)
            obj(26) = ReplaceNullValue(rw("end_dt_13"), dt.Columns("end_dt_13").DataType.ToString)
            obj(27) = ReplaceNullValue(rw("no_of_val_prd"), dt.Columns("no_of_val_prd").DataType.ToString)
            obj(28) = ReplaceNullValue(rw("def_prd_cd"), dt.Columns("def_prd_cd").DataType.ToString)
            obj(29) = ReplaceNullValue(rw("tmp_yr_cls_fg"), dt.Columns("tmp_yr_cls_fg").DataType.ToString)
            obj(30) = ReplaceNullValue(rw("warn_or_prev"), dt.Columns("warn_or_prev").DataType.ToString)
            obj(31) = ReplaceNullValue(rw("current_prd"), dt.Columns("current_prd").DataType.ToString)
            obj(32) = ReplaceNullValue(rw("wrnprv_glprd_fg_1"), dt.Columns("wrnprv_glprd_fg_1").DataType.ToString)
            obj(33) = ReplaceNullValue(rw("wrnprv_glprd_fg_2"), dt.Columns("wrnprv_glprd_fg_2").DataType.ToString)
            obj(34) = ReplaceNullValue(rw("wrnprv_glprd_fg_3"), dt.Columns("wrnprv_glprd_fg_3").DataType.ToString)
            obj(35) = ReplaceNullValue(rw("wrnprv_glprd_fg_4"), dt.Columns("wrnprv_glprd_fg_4").DataType.ToString)
            obj(36) = ReplaceNullValue(rw("wrnprv_glprd_fg_5"), dt.Columns("wrnprv_glprd_fg_5").DataType.ToString)
            obj(37) = ReplaceNullValue(rw("wrnprv_glprd_fg_6"), dt.Columns("wrnprv_glprd_fg_6").DataType.ToString)
            obj(38) = ReplaceNullValue(rw("wrnprv_glprd_fg_7"), dt.Columns("wrnprv_glprd_fg_7").DataType.ToString)
            obj(39) = ReplaceNullValue(rw("wrnprv_glprd_fg_8"), dt.Columns("wrnprv_glprd_fg_8").DataType.ToString)
            obj(40) = ReplaceNullValue(rw("wrnprv_glprd_fg_9"), dt.Columns("wrnprv_glprd_fg_9").DataType.ToString)
            obj(41) = ReplaceNullValue(rw("wrnprv_glprdfg_10"), dt.Columns("wrnprv_glprdfg_10").DataType.ToString)
            obj(42) = ReplaceNullValue(rw("wrnprv_glprdfg_11"), dt.Columns("wrnprv_glprdfg_11").DataType.ToString)
            obj(43) = ReplaceNullValue(rw("wrnprv_glprdfg_12"), dt.Columns("wrnprv_glprdfg_12").DataType.ToString)
            obj(44) = ReplaceNullValue(rw("wrnprv_glprdfg_13"), dt.Columns("wrnprv_glprdfg_13").DataType.ToString)
            obj(45) = ReplaceNullValue(rw("filler_0001"), dt.Columns("filler_0001").DataType.ToString)
            obj(46) = Convert.ToInt32(rw("str_dt_" & obj(31).ToString).ToString.Substring(0, 4))


        Next

        syprd = NewSystemPeriodSetting(
        obj(0).ToString,
        Convert.ToInt32(obj(1)),
        Convert.ToInt32(obj(2)),
        Convert.ToInt32(obj(3)),
        Convert.ToInt32(obj(4)),
        Convert.ToInt32(obj(5)),
        Convert.ToInt32(obj(6)),
        Convert.ToInt32(obj(7)),
        Convert.ToInt32(obj(8)),
        Convert.ToInt32(obj(9)),
        Convert.ToInt32(obj(10)),
        Convert.ToInt32(obj(11)),
        Convert.ToInt32(obj(12)),
        Convert.ToInt32(obj(13)),
        Convert.ToInt32(obj(14)),
        Convert.ToInt32(obj(15)),
        Convert.ToInt32(obj(16)),
        Convert.ToInt32(obj(17)),
        Convert.ToInt32(obj(18)),
        Convert.ToInt32(obj(19)),
        Convert.ToInt32(obj(20)),
        Convert.ToInt32(obj(21)),
        Convert.ToInt32(obj(22)),
        Convert.ToInt32(obj(23)),
        Convert.ToInt32(obj(24)),
        Convert.ToInt32(obj(25)),
        Convert.ToInt32(obj(26)),
        Convert.ToInt32(obj(27)),
        obj(28).ToString,
        obj(29).ToString,
        obj(30).ToString,
        Convert.ToInt32(obj(31)),
        obj(32).ToString,
        obj(33).ToString,
        obj(34).ToString,
        obj(35).ToString,
        obj(36).ToString,
        obj(37).ToString,
        obj(38).ToString,
        obj(39).ToString,
        obj(40).ToString,
        obj(41).ToString,
        obj(42).ToString,
        obj(43).ToString,
        obj(44).ToString,
        obj(45).ToString,
        obj(46).ToString)

        Return syprd

    End Function



    Private Function ReplaceNullValue(ByVal value As Object, ByVal datatype As String) As Object
        Dim val As Object = Nothing
        Select Case datatype
            Case "System.String"
                If value.Equals(DBNull.Value) Then
                    val = ""
                    Return Convert.ToString(val)
                Else
                    val = value
                    Return Convert.ToString(val)
                End If

            Case "System.Int16"

                If value Is DBNull.Value Then
                    val = 0
                    Return Convert.ToInt16(val)
                Else
                    val = value
                    Return Convert.ToInt16(val)
                End If

            Case "System.Int32"

                If value Is DBNull.Value Then
                    val = 0
                    Return Convert.ToInt32(val)
                Else
                    val = value
                    Return Convert.ToInt32(val)
                End If

            Case "System.Decimal"

                If value Is DBNull.Value Then
                    val = 0
                    Return Convert.ToDecimal(val)
                Else
                    val = value
                    Return Convert.ToDecimal(val)
                End If

            Case Else
                Return Nothing

        End Select
    End Function
End Class
