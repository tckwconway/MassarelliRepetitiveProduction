Imports System.Data.SqlClient


Public Class InventoryBin
    Public item_no As String
    Public item_filler As String
    Public loc As String
    Public bin_no As String
    Public bin_priority As Integer
    Public issue_priority As String
    Public issue_pri_tm As Integer
    Public bin_status As String
    Public prev_status As String
    Public received_dt As Integer
    Public contract_no As String
    Public unit_cost As Decimal
    Public orig_cost As Decimal
    Public qty_on_hand As Decimal
    Public qty_allocated As Decimal
    Public qty_bkord As Decimal
    Public qty_on_ord As Decimal
    Public cycle_count_cd As String
    Public last_count_dt As Integer
    Public tms_cntd_ytd As Integer
    Public pct_err_lst_cnt As Decimal
    Public frz_cost As Decimal
    Public frz_qty As Decimal
    Public frz_dt As Integer
    Public frz_tm As Integer
    Public max_designator As String
    Public max_value As Decimal
    Public max_uom As String
    Public max_batch_post As Decimal
    Public cube_width_uom As String
    Public cube_length_uom As String
    Public cube_height_uom As String
    Public cube_width As Decimal
    Public cube_length As Decimal
    Public cube_height As Decimal
    Public cube_qty_per As Decimal
    Public user_def_fld_1 As String
    Public user_def_fld_2 As String
    Public user_def_fld_3 As String
    Public user_def_fld_4 As String
    Public user_def_fld_5 As String
    Public tag_qty As Decimal
    Public filler_0003 As String

    Const TABLE As String = "IMINVBIN_SQL"


    Private Function NewInventoryBin(ByVal m_item_no As String, ByVal m_item_filler As String, ByVal m_loc As String, ByVal m_bin_no As String, ByVal m_bin_priority As Integer, _
                                            ByVal m_issue_priority As String, ByVal m_issue_pri_tm As Integer, ByVal m_bin_status As String, ByVal m_prev_status As String, ByVal m_received_dt As Integer, _
                                            ByVal m_contract_no As String, ByVal m_unit_cost As Decimal, ByVal m_orig_cost As Decimal, ByVal m_qty_on_hand As Decimal, ByVal m_qty_allocated As Decimal, _
                                            ByVal m_qty_bkord As Decimal, ByVal m_qty_on_ord As Decimal, ByVal m_cycle_count_cd As String, ByVal m_last_count_dt As Integer, ByVal m_tms_cntd_ytd As Integer, _
                                            ByVal m_pct_err_lst_cnt As Decimal, ByVal m_frz_cost As Decimal, ByVal m_frz_qty As Decimal, ByVal m_frz_dt As Integer, _
                                            ByVal m_frz_tm As Integer, ByVal m_max_designator As String, ByVal m_max_value As Decimal, ByVal m_max_uom As String, _
                                            ByVal m_max_batch_post As Decimal, ByVal m_cube_width_uom As String, ByVal m_cube_length_uom As String, ByVal m_cube_height_uom As String, _
                                            ByVal m_cube_width As Decimal, ByVal m_cube_length As Decimal, ByVal m_cube_height As Decimal, ByVal m_cube_qty_per As Decimal, _
                                            ByVal m_user_def_fld_1 As String, ByVal m_user_def_fld_2 As String, ByVal m_user_def_fld_3 As String, ByVal m_user_def_fld_4 As String, _
                                            ByVal m_user_def_fld_5 As String, ByVal m_tag_qty As Decimal, ByVal m_filler_0003 As String) As InventoryBin




        Dim invbin As New InventoryBin
        invbin.item_no = m_item_no
        invbin.item_filler = m_item_filler
        invbin.loc = m_loc
        invbin.bin_no = m_bin_no
        invbin.bin_priority = m_bin_priority
        invbin.issue_priority = m_issue_priority
        invbin.issue_pri_tm = m_issue_pri_tm
        invbin.bin_status = m_bin_status
        invbin.prev_status = m_prev_status
        invbin.received_dt = m_received_dt
        invbin.contract_no = m_contract_no
        invbin.unit_cost = m_unit_cost
        invbin.orig_cost = m_orig_cost
        invbin.qty_on_hand = m_qty_on_hand
        invbin.qty_allocated = m_qty_allocated
        invbin.qty_bkord = m_qty_bkord
        invbin.qty_on_ord = m_qty_on_ord
        invbin.cycle_count_cd = m_cycle_count_cd
        invbin.last_count_dt = m_last_count_dt
        invbin.tms_cntd_ytd = m_tms_cntd_ytd
        invbin.pct_err_lst_cnt = m_pct_err_lst_cnt
        invbin.frz_cost = m_frz_cost
        invbin.frz_qty = m_frz_qty
        invbin.frz_dt = m_frz_dt
        invbin.frz_tm = m_frz_tm
        invbin.max_designator = m_max_designator
        invbin.max_value = m_max_value
        invbin.max_uom = m_max_uom
        invbin.max_batch_post = m_max_batch_post
        invbin.cube_width_uom = m_cube_width_uom
        invbin.cube_length_uom = m_cube_length_uom
        invbin.cube_height_uom = m_cube_height_uom
        invbin.cube_width = m_cube_width
        invbin.cube_length = m_cube_length
        invbin.cube_height = m_cube_height
        invbin.cube_qty_per = m_cube_qty_per
        invbin.user_def_fld_1 = m_user_def_fld_1
        invbin.user_def_fld_2 = m_user_def_fld_2
        invbin.user_def_fld_3 = m_user_def_fld_3
        invbin.user_def_fld_4 = m_user_def_fld_4
        invbin.user_def_fld_5 = m_user_def_fld_5
        invbin.tag_qty = m_tag_qty
        invbin.filler_0003 = m_filler_0003

        Return invbin

    End Function

    Public Function NewInventoryBin(tranData As Object) As InventoryBin
        Dim invbin As New InventoryBin

        With invbin
            .item_no = tranData.ItemNo.ToString.Trim
            .item_filler = ""
            .loc = tranData.loc.ToString.Trim
            .bin_no = tranData.BinNo.ToString.Trim
            .bin_priority = 0
            .issue_priority = 0
            .issue_pri_tm = 0
            .bin_status = "A"
            .prev_status = Nothing
            .received_dt = cBusObj.GetMacolaDate(Now.Year.ToString, Format(Now.Month.ToString.PadLeft(2, "0")), Format(Now.Day.ToString.PadLeft(2, "0")))
            .contract_no = Nothing
            .unit_cost = 0
            .orig_cost = 0
            .qty_on_hand = tranData.NewBinQtyOnHand
            .qty_allocated = 0
            .qty_bkord = 0
            .qty_on_ord = 0
            .cycle_count_cd = Nothing
            .last_count_dt = 0
            .tms_cntd_ytd = 0
            .pct_err_lst_cnt = 0
            .frz_cost = 0
            .frz_qty = 0
            .frz_dt = 0
            .frz_tm = 0
            .max_designator = Nothing
            .max_value = 0
            .max_uom = Nothing
            .max_batch_post = 0
            .cube_width_uom = Nothing
            .cube_length_uom = Nothing
            .cube_height_uom = Nothing
            .cube_width = Nothing
            .cube_length = 0
            .cube_height = 0
            .cube_qty_per = 0
            .user_def_fld_1 = Nothing
            .user_def_fld_2 = Nothing
            .user_def_fld_3 = Nothing
            .user_def_fld_4 = Nothing
            .user_def_fld_5 = Nothing
            .tag_qty = 0
            .filler_0003 = Nothing
        End With
        Return invbin
    End Function
    Public Function NewInventoryBin() As InventoryBin
        Dim invbin As New InventoryBin

        With invbin
            .item_no = "" '
            .item_filler = "" '
            .loc = "" '
            .bin_no = "" '
            .bin_priority = 0 '
            .issue_priority = 0 '
            .issue_pri_tm = 0 '
            .bin_status = "A" '
            .prev_status = Nothing '
            .received_dt = 0 '
            .contract_no = Nothing '
            .unit_cost = 0 '
            .orig_cost = 0 '
            .qty_on_hand = 0 '
            .qty_allocated = 0 '
            .qty_bkord = 0 '
            .qty_on_ord = 0 '
            .cycle_count_cd = Nothing '
            .last_count_dt = 0 '
            .tms_cntd_ytd = 0 '
            .pct_err_lst_cnt = 0 '
            .frz_cost = 0 '
            .frz_qty = 0 '
            .frz_dt = 0 '
            .frz_tm = 0 '
            .max_designator = Nothing '
            .max_value = 0 '
            .max_uom = Nothing '
            .max_batch_post = 0 '
            .cube_width_uom = Nothing '
            .cube_length_uom = Nothing '
            .cube_height_uom = Nothing '
            .cube_width = Nothing '
            .cube_length = 0 '
            .cube_height = 0 '
            .cube_qty_per = 0 '
            .user_def_fld_1 = Nothing '
            .user_def_fld_2 = Nothing '
            .user_def_fld_3 = Nothing '
            .user_def_fld_4 = Nothing '
            .user_def_fld_5 = Nothing '
            .tag_qty = 0 '
            .filler_0003 = Nothing
        End With
        Return invbin
    End Function

    Public Function GetInventoryBin(ByVal item_no As String, bin_no As String, ByVal loc As String, ByVal cn As SqlConnection) As InventoryBin

        Dim invbin As New InventoryBin
        Dim dt As New DataTable
        Dim sSQL As String = "Select * from " & TABLE & " " _
                           & "Where item_no = '" & item_no & "' and bin_no = '" & bin_no & "' and loc = '" & loc & "' "

        dt = DAC.ExecuteSQL_DataTable(sSQL, cn, TABLE)

        invbin = PopulateInventoryBinSetting(dt)

        Return invbin

    End Function

    Public Sub Clear()

        item_no = ""
        item_filler = ""
        loc = ""
        bin_no = ""
        bin_priority = 0
        issue_priority = 0
        issue_pri_tm = 0
        bin_status = "A"
        prev_status = Nothing
        received_dt = 0
        contract_no = Nothing
        unit_cost = 0
        orig_cost = 0
        qty_on_hand = 0
        qty_allocated = 0
        qty_bkord = 0
        qty_on_ord = 0
        cycle_count_cd = Nothing
        last_count_dt = 0
        tms_cntd_ytd = 0
        pct_err_lst_cnt = 0
        frz_cost = 0
        frz_qty = 0
        frz_dt = 0
        frz_tm = 0
        max_designator = Nothing
        max_value = 0
        max_uom = Nothing
        max_batch_post = 0
        cube_width_uom = Nothing
        cube_length_uom = Nothing
        cube_height_uom = Nothing
        cube_width = Nothing
        cube_length = 0
        cube_height = 0
        cube_qty_per = 0
        user_def_fld_1 = Nothing
        user_def_fld_2 = Nothing
        user_def_fld_3 = Nothing
        user_def_fld_4 = Nothing
        user_def_fld_5 = Nothing
        tag_qty = 0
        filler_0003 = Nothing

    End Sub

    Private Function PopulateInventoryBinSetting(ByVal dt As DataTable) As InventoryBin
        Dim invbin As New InventoryBin
        Dim obj(42) As Object
        For Each rw As DataRow In dt.Rows

            obj(0) = ReplaceNullValue(rw("item_no"), dt.Columns("item_no").DataType.ToString)
            obj(1) = ReplaceNullValue(rw("item_filler"), dt.Columns("item_filler").DataType.ToString)
            obj(2) = ReplaceNullValue(rw("loc"), dt.Columns("loc").DataType.ToString)
            obj(3) = ReplaceNullValue(rw("bin_no"), dt.Columns("bin_no").DataType.ToString)
            obj(4) = ReplaceNullValue(rw("bin_priority"), dt.Columns("bin_priority").DataType.ToString)
            obj(5) = ReplaceNullValue(rw("issue_priority"), dt.Columns("issue_priority").DataType.ToString)
            obj(6) = ReplaceNullValue(rw("issue_pri_tm"), dt.Columns("issue_pri_tm").DataType.ToString)
            obj(7) = ReplaceNullValue(rw("bin_status"), dt.Columns("bin_status").DataType.ToString)
            obj(8) = ReplaceNullValue(rw("prev_status"), dt.Columns("prev_status").DataType.ToString)
            obj(9) = ReplaceNullValue(rw("received_dt"), dt.Columns("received_dt").DataType.ToString)
            obj(10) = ReplaceNullValue(rw("contract_no"), dt.Columns("contract_no").DataType.ToString)
            obj(11) = ReplaceNullValue(rw("unit_cost"), dt.Columns("unit_cost").DataType.ToString)
            obj(12) = ReplaceNullValue(rw("orig_cost"), dt.Columns("orig_cost").DataType.ToString)
            obj(13) = ReplaceNullValue(rw("qty_on_hand"), dt.Columns("qty_on_hand").DataType.ToString)
            obj(14) = ReplaceNullValue(rw("qty_allocated"), dt.Columns("qty_allocated").DataType.ToString)
            obj(15) = ReplaceNullValue(rw("qty_bkord"), dt.Columns("qty_bkord").DataType.ToString)
            obj(16) = ReplaceNullValue(rw("qty_on_ord"), dt.Columns("qty_on_ord").DataType.ToString)
            obj(17) = ReplaceNullValue(rw("cycle_count_cd"), dt.Columns("cycle_count_cd").DataType.ToString)
            obj(18) = ReplaceNullValue(rw("last_count_dt"), dt.Columns("last_count_dt").DataType.ToString)
            obj(19) = ReplaceNullValue(rw("tms_cntd_ytd"), dt.Columns("tms_cntd_ytd").DataType.ToString)
            obj(20) = ReplaceNullValue(rw("pct_err_lst_cnt"), dt.Columns("pct_err_lst_cnt").DataType.ToString)
            obj(21) = ReplaceNullValue(rw("frz_cost"), dt.Columns("frz_cost").DataType.ToString)
            obj(22) = ReplaceNullValue(rw("frz_qty"), dt.Columns("frz_qty").DataType.ToString)
            obj(23) = ReplaceNullValue(rw("frz_dt"), dt.Columns("frz_dt").DataType.ToString)
            obj(24) = ReplaceNullValue(rw("frz_tm"), dt.Columns("frz_tm").DataType.ToString)
            obj(25) = ReplaceNullValue(rw("max_designator"), dt.Columns("max_designator").DataType.ToString)
            obj(26) = ReplaceNullValue(rw("max_value"), dt.Columns("max_value").DataType.ToString)
            obj(27) = ReplaceNullValue(rw("max_uom"), dt.Columns("max_uom").DataType.ToString)
            obj(28) = ReplaceNullValue(rw("max_batch_post"), dt.Columns("max_batch_post").DataType.ToString)
            obj(29) = ReplaceNullValue(rw("cube_width_uom"), dt.Columns("cube_width_uom").DataType.ToString)
            obj(30) = ReplaceNullValue(rw("cube_length_uom"), dt.Columns("cube_length_uom").DataType.ToString)
            obj(31) = ReplaceNullValue(rw("cube_height_uom"), dt.Columns("cube_height_uom").DataType.ToString)
            obj(32) = ReplaceNullValue(rw("cube_width"), dt.Columns("cube_width").DataType.ToString)
            obj(33) = ReplaceNullValue(rw("cube_length"), dt.Columns("cube_length").DataType.ToString)
            obj(34) = ReplaceNullValue(rw("cube_height"), dt.Columns("cube_height").DataType.ToString)
            obj(35) = ReplaceNullValue(rw("cube_qty_per"), dt.Columns("cube_qty_per").DataType.ToString)
            obj(36) = ReplaceNullValue(rw("user_def_fld_1"), dt.Columns("user_def_fld_1").DataType.ToString)
            obj(37) = ReplaceNullValue(rw("user_def_fld_2"), dt.Columns("user_def_fld_2").DataType.ToString)
            obj(38) = ReplaceNullValue(rw("user_def_fld_3"), dt.Columns("user_def_fld_3").DataType.ToString)
            obj(39) = ReplaceNullValue(rw("user_def_fld_4"), dt.Columns("user_def_fld_4").DataType.ToString)
            obj(40) = ReplaceNullValue(rw("user_def_fld_5"), dt.Columns("user_def_fld_5").DataType.ToString)
            obj(41) = ReplaceNullValue(rw("tag_qty"), dt.Columns("tag_qty").DataType.ToString)
            obj(42) = ReplaceNullValue(rw("filler_0003"), dt.Columns("filler_0003").DataType.ToString)


        Next

        invbin = NewInventoryBin(
        obj(0).ToString,
        obj(1).ToString,
        obj(2).ToString,
        obj(3).ToString,
        Convert.ToInt32(obj(4)),
        obj(5).ToString,
        Convert.ToInt32(obj(6)),
        obj(7).ToString,
        obj(8).ToString,
        Convert.ToInt32(obj(9)),
        obj(10).ToString,
        Convert.ToDecimal(obj(11)),
        Convert.ToDecimal(obj(12)),
        Convert.ToDecimal(obj(13)),
        Convert.ToDecimal(obj(14)),
        Convert.ToDecimal(obj(15)),
        Convert.ToDecimal(obj(16)),
        obj(17).ToString,
        Convert.ToInt32(obj(18)),
        Convert.ToInt32(obj(19)),
        Convert.ToDecimal(obj(20)),
        Convert.ToDecimal(obj(21)),
        Convert.ToDecimal(obj(22)),
        Convert.ToInt32(obj(23)),
        Convert.ToInt32(obj(24)),
        obj(25).ToString,
        Convert.ToDecimal(obj(26)),
        obj(27).ToString,
        Convert.ToDecimal(obj(28)),
        obj(29).ToString,
        obj(30).ToString,
        obj(31).ToString,
        Convert.ToDecimal(obj(32)),
        Convert.ToDecimal(obj(33)),
        Convert.ToDecimal(obj(34)),
        Convert.ToDecimal(obj(35)),
        obj(36).ToString,
        obj(37).ToString,
        obj(38).ToString,
        obj(39).ToString,
        obj(40).ToString,
        Convert.ToDecimal(obj(41)),
        obj(42).ToString)

        Return invbin

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
