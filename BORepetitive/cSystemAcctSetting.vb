Imports System.Data.SqlClient

Public Class SystemAcctSetting


    Public mn_no As String
    Public sb_no As String
    Public dp_no As String
    Public acct_desc As String
    Public search_desc As String
    Public c_bal_mn_no As String
    Public c_bal_sb_no As String
    Public c_bal_dp_no As String
    Public consl_mn_no As String
    Public consl_sb_no As String
    Public consl_dp_no As String
    Public auto_dist_cd As Integer
    Public tb_sb_tot_lev As Integer
    Public fin_stmt_tp As String
    Public saf_tp As String
    Public par_ctl_cd As String
    Public comp_cd As String
    Public ratio_tp As String
    Public stat_fg As String
    Public valid_ap As String
    Public valid_ar As String
    Public valid_bb As String
    Public valid_bc As String
    Public valid_bm As String
    Public valid_cc As String
    Public valid_cm As String
    Public valid_cr As String
    Public valid_fx As String
    Public valid_gl As String
    Public valid_im As String
    Public valid_mc As String
    Public valid_lp As String
    Public valid_mr As String
    Public valid_ms As String
    Public valid_oe As String
    Public valid_po As String
    Public valid_pr As String
    Public valid_sc As String
    Public valid_sf As String
    Public valid_sr As String
    Public valid_pp As String
    Public valid_rm As String
    Public valid_de As String
    Public valid_je As String
    Public user_def_fld_1 As String
    Public user_def_fld_2 As String
    Public user_def_fld_3 As String
    Public user_def_fld_4 As String
    Public user_def_fld_5 As String
    Public use_id As String
    Public subtotal_desc As String
    Public inflation_type As String
    Public filler_0001 As String

    Const TABLE As String = "SYACTFIL_SQL"

    Private Function NewSystemAcctSetting(ByVal m_mn_no As String, ByVal m_sb_no As String, ByVal m_dp_no As String, ByVal m_acct_desc As String, ByVal m_search_desc As String, _
                   ByVal m_c_bal_mn_no As String, ByVal m_c_bal_sb_no As String, ByVal m_c_bal_dp_no As String, ByVal m_consl_mn_no As String, _
                   ByVal m_consl_sb_no As String, ByVal m_consl_dp_no As String, ByVal m_auto_dist_cd As Integer, ByVal m_tb_sb_tot_lev As Integer, _
                   ByVal m_fin_stmt_tp As String, ByVal m_saf_tp As String, ByVal m_par_ctl_cd As String, ByVal m_comp_cd As String, ByVal m_ratio_tp As String, _
                   ByVal m_stat_fg As String, ByVal m_valid_ap As String, ByVal m_valid_ar As String, ByVal m_valid_bb As String, ByVal m_valid_bc As String, _
                   ByVal m_valid_bm As String, ByVal m_valid_cc As String, ByVal m_valid_cm As String, ByVal m_valid_cr As String, ByVal m_valid_fx As String, _
                   ByVal m_valid_gl As String, ByVal m_valid_im As String, ByVal m_valid_mc As String, ByVal m_valid_lp As String, ByVal m_valid_mr As String, _
                   ByVal m_valid_ms As String, ByVal m_valid_oe As String, ByVal m_valid_po As String, ByVal m_valid_pr As String, ByVal m_valid_sc As String, _
                   ByVal m_valid_sf As String, ByVal m_valid_sr As String, ByVal m_valid_pp As String, ByVal m_valid_rm As String, ByVal m_valid_de As String, _
                   ByVal m_valid_je As String, ByVal m_user_def_fld_1 As String, ByVal m_user_def_fld_2 As String, ByVal m_user_def_fld_3 As String, _
                   ByVal m_user_def_fld_4 As String, ByVal m_user_def_fld_5 As String, ByVal m_use_id As String, ByVal m_subtotal_desc As String, _
                   ByVal m_inflation_type As String, ByVal m_filler_0001 As String) As SystemAcctSetting

        Dim syact As New SystemAcctSetting
        syact.mn_no = m_mn_no
        syact.sb_no = m_sb_no
        syact.dp_no = m_dp_no
        syact.acct_desc = m_acct_desc
        syact.search_desc = m_search_desc
        syact.c_bal_mn_no = m_c_bal_mn_no
        syact.c_bal_sb_no = m_c_bal_sb_no
        syact.c_bal_dp_no = m_c_bal_dp_no
        syact.consl_mn_no = m_consl_mn_no
        syact.consl_sb_no = m_consl_sb_no
        syact.consl_dp_no = m_consl_dp_no
        syact.auto_dist_cd = m_auto_dist_cd
        syact.tb_sb_tot_lev = m_tb_sb_tot_lev
        syact.fin_stmt_tp = m_fin_stmt_tp
        syact.saf_tp = m_saf_tp
        syact.par_ctl_cd = m_par_ctl_cd
        syact.comp_cd = m_comp_cd
        syact.ratio_tp = m_ratio_tp
        syact.stat_fg = m_stat_fg
        syact.valid_ap = m_valid_ap
        syact.valid_ar = m_valid_ar
        syact.valid_bb = m_valid_bb
        syact.valid_bc = m_valid_bc
        syact.valid_bm = m_valid_bm
        syact.valid_cc = m_valid_cc
        syact.valid_cm = m_valid_cm
        syact.valid_cr = m_valid_cr
        syact.valid_fx = m_valid_fx
        syact.valid_gl = m_valid_gl
        syact.valid_im = m_valid_im
        syact.valid_mc = m_valid_mc
        syact.valid_lp = m_valid_lp
        syact.valid_mr = m_valid_mr
        syact.valid_ms = m_valid_ms
        syact.valid_oe = m_valid_oe
        syact.valid_po = m_valid_po
        syact.valid_pr = m_valid_pr
        syact.valid_sc = m_valid_sc
        syact.valid_sf = m_valid_sf
        syact.valid_sr = m_valid_sr
        syact.valid_pp = m_valid_pp
        syact.valid_rm = m_valid_rm
        syact.valid_de = m_valid_de
        syact.valid_je = m_valid_je
        syact.user_def_fld_1 = m_user_def_fld_1
        syact.user_def_fld_2 = m_user_def_fld_2
        syact.user_def_fld_3 = m_user_def_fld_3
        syact.user_def_fld_4 = m_user_def_fld_4
        syact.user_def_fld_5 = m_user_def_fld_5
        syact.use_id = m_use_id
        syact.subtotal_desc = m_subtotal_desc
        syact.inflation_type = m_inflation_type
        syact.filler_0001 = m_filler_0001

        Return syact

    End Function

    Public Function GetSystemAcctSetting(ByVal mn_no As String, ByVal sb_no As String, ByVal db_no As String, ByVal cn As SqlConnection) As SystemAcctSetting
        ''mn_no, sb_no and db_no are the GL account numbering levels...
        ''INVENTORY is currently 10400000, 00000000 and 00000000 respectively

        Dim syact As New SystemAcctSetting
        Dim dt As New DataTable
        Dim sSQL As String = "Select * from " & TABLE & " " _
                           & "Where mn_no = '" & mn_no & "' " _
                           & "  and sb_no = '" & sb_no & "' " _
                           & "  and dp_no = '" & db_no & "' "

        dt = DAC.ExecuteSQL_DataTable(sSQL, cn, TABLE)

        syact = PopulateSystemAcctSettings(dt)

        Return syact

    End Function

    Private Function PopulateSystemAcctSettings(ByVal dt As DataTable) As SystemAcctSetting
        Dim syact As New SystemAcctSetting
        Dim obj(52) As Object
        For Each rw As DataRow In dt.Rows

            obj(0) = ReplaceNullValue(rw("mn_no"), dt.Columns("mn_no").DataType.ToString)
            obj(1) = ReplaceNullValue(rw("sb_no"), dt.Columns("sb_no").DataType.ToString)
            obj(2) = ReplaceNullValue(rw("dp_no"), dt.Columns("dp_no").DataType.ToString)
            obj(3) = ReplaceNullValue(rw("acct_desc"), dt.Columns("acct_desc").DataType.ToString)
            obj(4) = ReplaceNullValue(rw("search_desc"), dt.Columns("search_desc").DataType.ToString)
            obj(5) = ReplaceNullValue(rw("c_bal_mn_no"), dt.Columns("c_bal_mn_no").DataType.ToString)
            obj(6) = ReplaceNullValue(rw("c_bal_sb_no"), dt.Columns("c_bal_sb_no").DataType.ToString)
            obj(7) = ReplaceNullValue(rw("c_bal_dp_no"), dt.Columns("c_bal_dp_no").DataType.ToString)
            obj(8) = ReplaceNullValue(rw("consl_mn_no"), dt.Columns("consl_mn_no").DataType.ToString)
            obj(9) = ReplaceNullValue(rw("consl_sb_no"), dt.Columns("consl_sb_no").DataType.ToString)
            obj(10) = ReplaceNullValue(rw("consl_dp_no"), dt.Columns("consl_dp_no").DataType.ToString)
            obj(11) = ReplaceNullValue(rw("auto_dist_cd"), dt.Columns("auto_dist_cd").DataType.ToString)
            obj(12) = ReplaceNullValue(rw("tb_sb_tot_lev"), dt.Columns("tb_sb_tot_lev").DataType.ToString)
            obj(13) = ReplaceNullValue(rw("fin_stmt_tp"), dt.Columns("fin_stmt_tp").DataType.ToString)
            obj(14) = ReplaceNullValue(rw("saf_tp"), dt.Columns("saf_tp").DataType.ToString)
            obj(15) = ReplaceNullValue(rw("par_ctl_cd"), dt.Columns("par_ctl_cd").DataType.ToString)
            obj(16) = ReplaceNullValue(rw("comp_cd"), dt.Columns("comp_cd").DataType.ToString)
            obj(17) = ReplaceNullValue(rw("ratio_tp"), dt.Columns("ratio_tp").DataType.ToString)
            obj(18) = ReplaceNullValue(rw("stat_fg"), dt.Columns("stat_fg").DataType.ToString)
            obj(19) = ReplaceNullValue(rw("valid_ap"), dt.Columns("valid_ap").DataType.ToString)
            obj(20) = ReplaceNullValue(rw("valid_ar"), dt.Columns("valid_ar").DataType.ToString)
            obj(21) = ReplaceNullValue(rw("valid_bb"), dt.Columns("valid_bb").DataType.ToString)
            obj(22) = ReplaceNullValue(rw("valid_bc"), dt.Columns("valid_bc").DataType.ToString)
            obj(23) = ReplaceNullValue(rw("valid_bm"), dt.Columns("valid_bm").DataType.ToString)
            obj(24) = ReplaceNullValue(rw("valid_cc"), dt.Columns("valid_cc").DataType.ToString)
            obj(25) = ReplaceNullValue(rw("valid_cm"), dt.Columns("valid_cm").DataType.ToString)
            obj(26) = ReplaceNullValue(rw("valid_cr"), dt.Columns("valid_cr").DataType.ToString)
            obj(27) = ReplaceNullValue(rw("valid_fx"), dt.Columns("valid_fx").DataType.ToString)
            obj(28) = ReplaceNullValue(rw("valid_gl"), dt.Columns("valid_gl").DataType.ToString)
            obj(29) = ReplaceNullValue(rw("valid_im"), dt.Columns("valid_im").DataType.ToString)
            obj(30) = ReplaceNullValue(rw("valid_mc"), dt.Columns("valid_mc").DataType.ToString)
            obj(31) = ReplaceNullValue(rw("valid_lp"), dt.Columns("valid_lp").DataType.ToString)
            obj(32) = ReplaceNullValue(rw("valid_mr"), dt.Columns("valid_mr").DataType.ToString)
            obj(33) = ReplaceNullValue(rw("valid_ms"), dt.Columns("valid_ms").DataType.ToString)
            obj(34) = ReplaceNullValue(rw("valid_oe"), dt.Columns("valid_oe").DataType.ToString)
            obj(35) = ReplaceNullValue(rw("valid_po"), dt.Columns("valid_po").DataType.ToString)
            obj(36) = ReplaceNullValue(rw("valid_pr"), dt.Columns("valid_pr").DataType.ToString)
            obj(37) = ReplaceNullValue(rw("valid_sc"), dt.Columns("valid_sc").DataType.ToString)
            obj(38) = ReplaceNullValue(rw("valid_sf"), dt.Columns("valid_sf").DataType.ToString)
            obj(39) = ReplaceNullValue(rw("valid_sr"), dt.Columns("valid_sr").DataType.ToString)
            obj(40) = ReplaceNullValue(rw("valid_pp"), dt.Columns("valid_pp").DataType.ToString)
            obj(41) = ReplaceNullValue(rw("valid_rm"), dt.Columns("valid_rm").DataType.ToString)
            obj(42) = ReplaceNullValue(rw("valid_de"), dt.Columns("valid_de").DataType.ToString)
            obj(43) = ReplaceNullValue(rw("valid_je"), dt.Columns("valid_je").DataType.ToString)
            obj(44) = ReplaceNullValue(rw("user_def_fld_1"), dt.Columns("user_def_fld_1").DataType.ToString)
            obj(45) = ReplaceNullValue(rw("user_def_fld_2"), dt.Columns("user_def_fld_2").DataType.ToString)
            obj(46) = ReplaceNullValue(rw("user_def_fld_3"), dt.Columns("user_def_fld_3").DataType.ToString)
            obj(47) = ReplaceNullValue(rw("user_def_fld_4"), dt.Columns("user_def_fld_4").DataType.ToString)
            obj(48) = ReplaceNullValue(rw("user_def_fld_5"), dt.Columns("user_def_fld_5").DataType.ToString)
            obj(49) = ReplaceNullValue(rw("use_id"), dt.Columns("use_id").DataType.ToString)
            obj(50) = ReplaceNullValue(rw("subtotal_desc"), dt.Columns("subtotal_desc").DataType.ToString)
            obj(51) = ReplaceNullValue(rw("inflation_type"), dt.Columns("inflation_type").DataType.ToString)
            obj(52) = ReplaceNullValue(rw("filler_0001"), dt.Columns("filler_0001").DataType.ToString)

            syact = NewSystemAcctSetting(obj(0).ToString,
                                        obj(1).ToString,
                                        obj(2).ToString,
                                        obj(3).ToString,
                                        obj(4).ToString,
                                        obj(5).ToString,
                                        obj(6).ToString,
                                        obj(7).ToString,
                                        obj(8).ToString,
                                        obj(9).ToString,
                                        obj(10).ToString,
                                        Convert.ToInt32(obj(11)),
                                        Convert.ToInt32(obj(12)),
                                        obj(13).ToString,
                                        obj(14).ToString,
                                        obj(15).ToString,
                                        obj(16).ToString,
                                        obj(17).ToString,
                                        obj(18).ToString,
                                        obj(19).ToString,
                                        obj(20).ToString,
                                        obj(21).ToString,
                                        obj(22).ToString,
                                        obj(23).ToString,
                                        obj(24).ToString,
                                        obj(25).ToString,
                                        obj(26).ToString,
                                        obj(27).ToString,
                                        obj(28).ToString,
                                        obj(29).ToString,
                                        obj(30).ToString,
                                        obj(31).ToString,
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
                                        obj(46).ToString,
                                        obj(47).ToString,
                                        obj(48).ToString,
                                        obj(49).ToString,
                                        obj(50).ToString,
                                        obj(51).ToString,
                                        obj(52).ToString)

        Next

        Return syact
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

        End Select

        Return val

    End Function
End Class


'    'Public Class Get_DataRow
'        Implements IDisposable
'        Friend row As DataRow
'        'Public Class FA_DataRow

'        Dim sErrMsg As String = ""

'        Sub New(Optional ByRef UseThisRow As DataRow = Nothing)

'            row = UseThisRow

'        End Sub

'        Sub Add(ByVal UseThisRow As DataRow)

'            row = UseThisRow

'        End Sub

'        Sub Clear()

'            row = Nothing

'        End Sub

'        Public ReadOnly Property ToDataRow() As DataRow

'            Get

'                Return row

'            End Get

'        End Property

'        ReadOnly Property HasData() As Boolean

'            Get

'                Return IIf(Not row Is Nothing, True, False)

'            End Get

'        End Property

'        ReadOnly Property Count() As Integer

'            Get

'                Dim value As Integer = 0

'                If Not row Is Nothing Then

'                    If Not row.Table Is Nothing Then

'                        value = row.Table.Columns.Count

'                    End If

'                End If

'                Return value

'            End Get

'        End Property

'        ''' <summary>

'        ''' Return a list of column names from the record.

'        ''' </summary>

'        ''' <value></value>

'        ''' <returns></returns>

'        ''' <remarks></remarks>

'        Public ReadOnly Property ColumnList() As String()

'            Get

'                Dim sValues() As String = Nothing

'                Dim n As Integer = 0

'                If row IsNot Nothing Then

'                    If row.Table IsNot Nothing Then

'                        ReDim sValues(0 To row.Table.Columns.Count - 1)

'                        For Each col As DataColumn In row.Table.Columns

'                            sValues(n) = col.ColumnName

'                            n += 1

'                        Next

'                    End If

'                End If

'                Return sValues

'            End Get

'        End Property

'        ''' <summary>

'        ''' Return a list of values from the record.

'        ''' </summary>

'        ''' <value></value>

'        ''' <returns></returns>

'        ''' <remarks></remarks>

'        Public ReadOnly Property ValueList() As String()

'            Get

'                Dim sValues() As String = Nothing

'                Dim n As Integer = 0

'                If row IsNot Nothing Then

'                    ReDim sValues(0 To row.Table.Columns.Count - 1)

'                    For n = 0 To row.Table.Columns.Count

'                        sValues(n) = row.Item(n).ToString()

'                        n += 1

'                    Next

'                End If

'                Return sValues

'            End Get

'        End Property

'        ''' <summary>

'        ''' Get the value of the 1st item in the record.

'        ''' </summary>

'        ''' <value></value>

'        ''' <returns></returns>

'        ''' <remarks></remarks>

'        Public ReadOnly Property FirstItem() As Object

'            Get

'                If HasData Then

'                    If row.Table.Columns.Count > 0 Then

'                        Return row.Item(0)

'                    End If

'                End If

'                Return Nothing

'            End Get

'        End Property

'        ''' <summary>

'        ''' Get the value of the last item in the record.

'        ''' </summary>

'        ''' <value></value>

'        ''' <returns></returns>

'        ''' <remarks></remarks>

'        Public ReadOnly Property LastItem() As Object

'            Get

'                If HasData Then

'                    If row.Table.Columns.Count > 0 Then

'                        Return row.Item(row.Table.Columns.Count - 1)

'                    End If

'                End If

'                Return Nothing

'            End Get

'        End Property

'        ''' <summary>

'        ''' Get the value of the field specified at this index in the record.

'        ''' </summary>

'        ''' <value></value>

'        ''' <returns></returns>

'        ''' <remarks></remarks>

'        Public Property Item(ByVal Index As Integer) As Object

'            Get

'                If HasData Then

'                    If Index >= 0 And Index < row.Table.Columns.Count Then

'                        Return row.Item(Index)

'                    End If

'                End If

'                Return Nothing

'            End Get

'            Set(ByVal value As Object)

'                If HasData Then

'                    If Index >= 0 And Index < row.Table.Columns.Count Then

'                        row.Item(Index) = value

'                    End If

'                End If

'            End Set

'        End Property

'        ''' <summary>

'        ''' Get the value of the field specified by this column name in the record.

'        ''' </summary>

'        ''' <value></value>

'        ''' <returns></returns>

'        ''' <remarks></remarks>

'        Public Property Item(ByVal sColName As String) As Object

'            Get

'                If HasData AndAlso Contains(sColName) Then

'                    Return row.Item(sColName)

'                End If

'                Return Nothing

'            End Get

'            Set(ByVal value As Object)

'                If HasData AndAlso Contains(sColName) Then

'                    row.Item(sColName) = value

'                End If

'            End Set

'        End Property

'        Shadows ReadOnly Property ToString(ByVal sColName As String, Optional ByVal bTrim As Boolean = True) As String

'            Get

'                Dim sColData As String = ""

'                If HasData AndAlso Contains(sColName) Then

'                    Try

'                        If Contains(sColName) Then

'                            sColData = row.Item(sColName).ToString

'                        End If

'                    Catch ex As Exception

'                        sErrMsg = ex.Message

'                    End Try

'                End If

'    End Function

'#Region "IDisposable Support"
'                Private disposedValue As Boolean ' To detect redundant calls

'                ' IDisposable
'    Protected     Overridable     Sub Dispose(ByVal disposing As Boolean)
'            If Not Me.disposedValue Then
'                If disposing Then
'                    ' TODO: dispose managed state (managed objects).
'                End If

'                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
'                ' TODO: set large fields to null.
'            End If
'            Me.disposedValue = True
'        End Sub

'        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
'        'Protected Overrides Sub Finalize()
'        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
'        '    Dispose(False)
'        '    MyBase.Finalize()
'        'End Sub

'        ' This code added by Visual Basic to correctly implement the disposable pattern.
'        Public Sub Dispose() Implements IDisposable.Dispose
'            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
'            Dispose(True)
'            GC.SuppressFinalize(Me)
'        End Sub
'#End Region

'    End Class
