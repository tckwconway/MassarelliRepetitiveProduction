Imports System.ComponentModel
Imports System.Data.DataRowView
Imports System.Data.SqlClient
Imports System.Data
Imports System.Text
Imports System.Xml
Imports System.IO
Imports System.Data.SqlTypes


Public Class InventoryLocation
    Public item_no As String
    Public item_filler As String
    Public loc As String
    Public status As String
    Public prev_status As String
    Public mult_bin_fg As String
    Public qty_on_hand As Decimal
    Public qty_allocated As Decimal
    Public qty_bkord As Decimal
    Public qty_on_ord As Decimal
    Public reorder_lvl As Decimal
    Public ord_up_to_lvl As Decimal
    Public price As Decimal
    Public avg_cost As Decimal
    Public last_cost As Decimal
    Public std_cost As Decimal
    Public prcs_apply_fg As String
    Public discs_apply_fg As String
    Public starting_sls_dt As Integer
    Public ending_sls_dt As Integer
    Public last_sold_dt As Integer
    Public sls_price As Decimal
    Public qty_last_sold As Decimal
    Public cycle_count_cd As String
    Public last_count_dt As Integer
    Public tms_cntd_ytd As Integer
    Public pct_err_last_cnt As Decimal
    Public frz_cost As Decimal
    Public frz_qty As Decimal
    Public frz_dt As Integer
    Public frz_tm As Integer
    Public usage_ptd As Decimal
    Public qty_sld_ptd As Decimal
    Public qty_scrp_ptd As Decimal
    Public sls_ptd As Decimal
    Public cost_ptd As Decimal
    Public usage_ytd As Decimal
    Public qty_sold_ytd As Decimal
    Public qty_scrp_ytd As Decimal
    Public qty_returned_ytd As Decimal
    Public sls_ytd As Decimal
    Public cost_ytd As Decimal
    Public prior_year_usage As Decimal
    Public qty_sold_last_yr As Decimal
    Public qty_scrp_last_yr As Decimal
    Public prior_year_sls As Decimal
    Public cost_last_yr As Decimal
    Public recom_min_ord As Decimal
    Public economic_ord_qty As Decimal
    Public avg_usage As Decimal
    Public po_lead_tm As Integer
    Public byr_plnr As String
    Public doc_to_stk_ld_tm As Integer
    Public rollup_prc As String
    Public target_margin As Integer
    Public inv_class As String
    Public po_min As Decimal
    Public po_max As Decimal
    Public safety_stk As Decimal
    Public avg_frcst_error As Decimal
    Public sum_of_errors As Decimal
    Public usg_wght_fctr As Decimal
    Public safety_fctr As Decimal
    Public usage_filter As Decimal
    Public po_mult As Integer
    Public active_ords As Integer
    Public vend_no As String
    Public tax_sched As String
    Public prod_cat As String
    Public picking_seq As String
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
    Public landed_cost_cd As String
    Public landed_cost_cd_2 As String
    Public landed_cost_cd_3 As String
    Public landed_cost_cd_4 As String
    Public landed_cost_cd_5 As String
    Public landed_cost_cd_6 As String
    Public landed_cost_cd_7 As String
    Public landed_cost_cd_8 As String
    Public landed_cost_cd_9 As String
    Public landed_cost_cd_10 As String
    Public loc_qty_fld As Decimal
    Public tag_qty As Decimal
    Public tag_cost As Decimal
    Public tag_frz_dt As Integer
    Public filler_0002 As String

    Const TABLE As String = "IMINVLOC_SQL"

    Private Function NewInventoryLocation(ByVal mitem_no As String, ByVal mitem_filler As String, ByVal mloc As String, ByVal mstatus As String, ByVal mprev_status As String, _
                                          ByVal mmult_bin_fg As String, ByVal mqty_on_hand As Decimal, ByVal mqty_allocated As Decimal, ByVal mqty_bkord As Decimal, _
                                          ByVal mqty_on_ord As Decimal, ByVal mreorder_lvl As Decimal, ByVal mord_up_to_lvl As Decimal, ByVal mprice As Decimal, _
                                          ByVal mavg_cost As Decimal, ByVal mlast_cost As Decimal, ByVal mstd_cost As Decimal, ByVal mprcs_apply_fg As String, _
                                          ByVal mdiscs_apply_fg As String, ByVal mstarting_sls_dt As Integer, ByVal mending_sls_dt As Integer, ByVal mlast_sold_dt As Integer, _
                                          ByVal msls_price As Decimal, ByVal mqty_last_sold As Decimal, ByVal mcycle_count_cd As String, ByVal mlast_count_dt As Integer, _
                                          ByVal mtms_cntd_ytd As Integer, ByVal mpct_err_last_cnt As Decimal, ByVal mfrz_cost As Decimal, ByVal mfrz_qty As Decimal, _
                                          ByVal mfrz_dt As Integer, ByVal mfrz_tm As Integer, ByVal musage_ptd As Decimal, ByVal mqty_sld_ptd As Decimal, _
                                          ByVal mqty_scrp_ptd As Decimal, ByVal msls_ptd As Decimal, ByVal mcost_ptd As Decimal, ByVal musage_ytd As Decimal, _
                                          ByVal mqty_sold_ytd As Decimal, ByVal mqty_scrp_ytd As Decimal, ByVal mqty_returned_ytd As Decimal, ByVal msls_ytd As Decimal, _
                                          ByVal mcost_ytd As Decimal, ByVal mprior_year_usage As Decimal, ByVal mqty_sold_last_yr As Decimal, ByVal mqty_scrp_last_yr As Decimal, _
                                          ByVal mprior_year_sls As Decimal, ByVal mcost_last_yr As Decimal, ByVal mrecom_min_ord As Decimal, ByVal meconomic_ord_qty As Decimal, _
                                          ByVal mavg_usage As Decimal, ByVal mpo_lead_tm As Integer, ByVal mbyr_plnr As String, ByVal mdoc_to_stk_ld_tm As Integer, _
                                          ByVal mrollup_prc As String, ByVal mtarget_margin As Integer, ByVal minv_class As String, ByVal mpo_min As Decimal, _
                                          ByVal mpo_max As Decimal, ByVal msafety_stk As Decimal, ByVal mavg_frcst_error As Decimal, _
                                          ByVal msum_of_errors As Decimal, ByVal musg_wght_fctr As Decimal, ByVal msafety_fctr As Decimal, _
                                          ByVal musage_filter As Decimal, ByVal mpo_mult As Integer, ByVal mactive_ords As Integer, ByVal mvend_no As String, _
                                          ByVal mtax_sched As String, ByVal mprod_cat As String, ByVal mpicking_seq As String, ByVal mcube_width_uom As String, _
                                          ByVal mcube_length_uom As String, ByVal mcube_height_uom As String, ByVal mcube_width As Decimal, ByVal mcube_length As Decimal, _
                                          ByVal mcube_height As Decimal, ByVal mcube_qty_per As Decimal, ByVal muser_def_fld_1 As String, ByVal muser_def_fld_2 As String, _
                                          ByVal muser_def_fld_3 As String, ByVal muser_def_fld_4 As String, ByVal muser_def_fld_5 As String, ByVal mlanded_cost_cd As String, _
                                          ByVal mlanded_cost_cd_2 As String, ByVal mlanded_cost_cd_3 As String, ByVal mlanded_cost_cd_4 As String, ByVal mlanded_cost_cd_5 As String, _
                                          ByVal mlanded_cost_cd_6 As String, ByVal mlanded_cost_cd_7 As String, ByVal mlanded_cost_cd_8 As String, ByVal mlanded_cost_cd_9 As String, _
                                          ByVal mlanded_cost_cd_10 As String, ByVal mloc_qty_fld As Decimal, ByVal mtag_qty As Decimal, ByVal mtag_cost As Decimal, _
                                          ByVal mtag_frz_dt As Integer, ByVal mfiller_0002 As String) As InventoryLocation

        Dim invloc As New InventoryLocation
        invloc.item_no = item_no
        invloc.item_filler = mitem_filler
        invloc.loc = mloc
        invloc.status = mstatus
        invloc.prev_status = mprev_status
        invloc.mult_bin_fg = mmult_bin_fg
        invloc.qty_on_hand = mqty_on_hand
        invloc.qty_allocated = mqty_allocated
        invloc.qty_bkord = mqty_bkord
        invloc.qty_on_ord = mqty_on_ord
        invloc.reorder_lvl = mreorder_lvl
        invloc.ord_up_to_lvl = mord_up_to_lvl
        invloc.price = mprice
        invloc.avg_cost = mavg_cost
        invloc.last_cost = mlast_cost
        invloc.std_cost = mstd_cost
        invloc.prcs_apply_fg = mprcs_apply_fg
        invloc.discs_apply_fg = mdiscs_apply_fg
        invloc.starting_sls_dt = mstarting_sls_dt
        invloc.ending_sls_dt = mending_sls_dt
        invloc.last_sold_dt = mlast_sold_dt
        invloc.sls_price = msls_price
        invloc.qty_last_sold = mqty_last_sold
        invloc.cycle_count_cd = mcycle_count_cd
        invloc.last_count_dt = mlast_count_dt
        invloc.tms_cntd_ytd = mtms_cntd_ytd
        invloc.pct_err_last_cnt = mpct_err_last_cnt
        invloc.frz_cost = mfrz_cost
        invloc.frz_qty = mfrz_qty
        invloc.frz_dt = mfrz_dt
        invloc.frz_tm = mfrz_tm
        invloc.usage_ptd = musage_ptd
        invloc.qty_sld_ptd = mqty_sld_ptd
        invloc.qty_scrp_ptd = mqty_scrp_ptd
        invloc.sls_ptd = msls_ptd
        invloc.cost_ptd = mcost_ptd
        invloc.usage_ytd = musage_ytd
        invloc.qty_sold_ytd = mqty_sold_ytd
        invloc.qty_scrp_ytd = mqty_scrp_ytd
        invloc.qty_returned_ytd = mqty_returned_ytd
        invloc.sls_ytd = msls_ytd
        invloc.cost_ytd = mcost_ytd
        invloc.prior_year_usage = mprior_year_usage
        invloc.qty_sold_last_yr = mqty_sold_last_yr
        invloc.qty_scrp_last_yr = mqty_scrp_last_yr
        invloc.prior_year_sls = mprior_year_sls
        invloc.cost_last_yr = mcost_last_yr
        invloc.recom_min_ord = mrecom_min_ord
        invloc.economic_ord_qty = meconomic_ord_qty
        invloc.avg_usage = mavg_usage
        invloc.po_lead_tm = mpo_lead_tm
        invloc.byr_plnr = mbyr_plnr
        invloc.doc_to_stk_ld_tm = mdoc_to_stk_ld_tm
        invloc.rollup_prc = mrollup_prc
        invloc.target_margin = mtarget_margin
        invloc.inv_class = minv_class
        invloc.po_min = mpo_min
        invloc.po_max = mpo_max
        invloc.safety_stk = msafety_stk
        invloc.avg_frcst_error = mavg_frcst_error
        invloc.sum_of_errors = msum_of_errors
        invloc.usg_wght_fctr = musg_wght_fctr
        invloc.safety_fctr = msafety_fctr
        invloc.usage_filter = musage_filter
        invloc.po_mult = mpo_mult
        invloc.active_ords = mactive_ords
        invloc.vend_no = mvend_no
        invloc.tax_sched = mtax_sched
        invloc.prod_cat = mprod_cat
        invloc.picking_seq = mpicking_seq
        invloc.cube_width_uom = mcube_width_uom
        invloc.cube_length_uom = mcube_length_uom
        invloc.cube_height_uom = mcube_height_uom
        invloc.cube_width = mcube_width
        invloc.cube_length = mcube_length
        invloc.cube_height = mcube_height
        invloc.cube_qty_per = mcube_qty_per
        invloc.user_def_fld_1 = muser_def_fld_1
        invloc.user_def_fld_2 = muser_def_fld_2
        invloc.user_def_fld_3 = muser_def_fld_3
        invloc.user_def_fld_4 = muser_def_fld_4
        invloc.user_def_fld_5 = muser_def_fld_5
        invloc.landed_cost_cd = mlanded_cost_cd
        invloc.landed_cost_cd_2 = mlanded_cost_cd_2
        invloc.landed_cost_cd_3 = mlanded_cost_cd_3
        invloc.landed_cost_cd_4 = mlanded_cost_cd_4
        invloc.landed_cost_cd_5 = mlanded_cost_cd_5
        invloc.landed_cost_cd_6 = mlanded_cost_cd_6
        invloc.landed_cost_cd_7 = mlanded_cost_cd_7
        invloc.landed_cost_cd_8 = mlanded_cost_cd_8
        invloc.landed_cost_cd_9 = mlanded_cost_cd_9
        invloc.landed_cost_cd_10 = mlanded_cost_cd_10
        invloc.loc_qty_fld = mloc_qty_fld
        invloc.tag_qty = mtag_qty
        invloc.tag_cost = mtag_cost
        invloc.tag_frz_dt = mtag_frz_dt
        invloc.filler_0002 = mfiller_0002

        Return invloc
    End Function

    Public Function NewInventoryLocation() As InventoryLocation
        Dim invloc As New InventoryLocation

        invloc.item_no = ""
        invloc.item_filler = ""
        invloc.loc = ""
        invloc.status = ""
        invloc.prev_status = ""
        invloc.mult_bin_fg = ""
        invloc.qty_on_hand = 0
        invloc.qty_allocated = 0
        invloc.qty_bkord = 0
        invloc.qty_on_ord = 0
        invloc.reorder_lvl = 0
        invloc.ord_up_to_lvl = 0
        invloc.price = 0
        invloc.avg_cost = 0
        invloc.last_cost = 0
        invloc.std_cost = 0
        invloc.prcs_apply_fg = ""
        invloc.discs_apply_fg = ""
        invloc.starting_sls_dt = 0
        invloc.ending_sls_dt = 0
        invloc.last_sold_dt = 0
        invloc.sls_price = 0
        invloc.qty_last_sold = 0
        invloc.cycle_count_cd = ""
        invloc.last_count_dt = 0
        invloc.tms_cntd_ytd = 0
        invloc.pct_err_last_cnt = 0
        invloc.frz_cost = 0
        invloc.frz_qty = 0
        invloc.frz_dt = 0
        invloc.frz_tm = 0
        invloc.usage_ptd = 0
        invloc.qty_sld_ptd = 0
        invloc.qty_scrp_ptd = 0
        invloc.sls_ptd = 0
        invloc.cost_ptd = 0
        invloc.usage_ytd = 0
        invloc.qty_sold_ytd = 0
        invloc.qty_scrp_ytd = 0
        invloc.qty_returned_ytd = 0
        invloc.sls_ytd = 0
        invloc.cost_ytd = 0
        invloc.prior_year_usage = 0
        invloc.qty_sold_last_yr = 0
        invloc.qty_scrp_last_yr = 0
        invloc.prior_year_sls = 0
        invloc.cost_last_yr = 0
        invloc.recom_min_ord = 0
        invloc.economic_ord_qty = 0
        invloc.avg_usage = 0
        invloc.po_lead_tm = 0
        invloc.byr_plnr = ""
        invloc.doc_to_stk_ld_tm = 0
        invloc.rollup_prc = ""
        invloc.target_margin = 0
        invloc.inv_class = ""
        invloc.po_min = 0
        invloc.po_max = 0
        invloc.safety_stk = 0
        invloc.avg_frcst_error = 0
        invloc.sum_of_errors = 0
        invloc.usg_wght_fctr = 0
        invloc.safety_fctr = 0
        invloc.usage_filter = 0
        invloc.po_mult = 0
        invloc.active_ords = 0
        invloc.vend_no = ""
        invloc.tax_sched = ""
        invloc.prod_cat = ""
        invloc.picking_seq = ""
        invloc.cube_width_uom = ""
        invloc.cube_length_uom = ""
        invloc.cube_height_uom = ""
        invloc.cube_width = 0
        invloc.cube_length = 0
        invloc.cube_height = 0
        invloc.cube_qty_per = 0
        invloc.user_def_fld_1 = ""
        invloc.user_def_fld_2 = ""
        invloc.user_def_fld_3 = ""
        invloc.user_def_fld_4 = ""
        invloc.user_def_fld_5 = ""
        invloc.landed_cost_cd = ""
        invloc.landed_cost_cd_2 = ""
        invloc.landed_cost_cd_3 = ""
        invloc.landed_cost_cd_4 = ""
        invloc.landed_cost_cd_5 = ""
        invloc.landed_cost_cd_6 = ""
        invloc.landed_cost_cd_7 = ""
        invloc.landed_cost_cd_8 = ""
        invloc.landed_cost_cd_9 = ""
        invloc.landed_cost_cd_10 = ""
        invloc.loc_qty_fld = 0
        invloc.tag_qty = 0
        invloc.tag_cost = 0
        invloc.tag_frz_dt = 0
        invloc.filler_0002 = ""

        Return invloc

    End Function



    Public Function GetInventoryLocation(ByVal item_no As String, ByVal loc As String, ByVal cn As SqlConnection) As InventoryLocation

        Dim invloc As New InventoryLocation
        Dim dt As New DataTable
        Dim sSQL As String = "Select * from " & TABLE & " " _
                           & "Where item_no = '" & item_no & "' and loc = '" & loc & "' "

        dt = DAC.ExecuteSQL_DataTable(sSQL, cn, TABLE)

        invloc = PopulateInventoryLocation(dt)

        Return invloc

    End Function




    Private Function PopulateInventoryLocation(ByVal dt As DataTable) As InventoryLocation
        Dim invloc As New InventoryLocation
        Dim obj(96) As Object
        For Each rw As DataRow In dt.Rows

            obj(0) = ReplaceNullValue(rw("item_no").ToString.Trim, dt.Columns("item_no").DataType.ToString.Trim.Trim)
            obj(1) = ReplaceNullValue(rw("item_filler").ToString.Trim, dt.Columns("item_filler").DataType.ToString.Trim)
            obj(2) = ReplaceNullValue(rw("loc").ToString.Trim, dt.Columns("loc").DataType.ToString.Trim)
            obj(3) = ReplaceNullValue(rw("status").ToString.Trim, dt.Columns("status").DataType.ToString.Trim)
            obj(4) = ReplaceNullValue(rw("prev_status").ToString.Trim, dt.Columns("prev_status").DataType.ToString.Trim)
            obj(5) = ReplaceNullValue(rw("mult_bin_fg").ToString.Trim, dt.Columns("mult_bin_fg").DataType.ToString.Trim)
            obj(6) = ReplaceNullValue(rw("qty_on_hand").ToString.Trim, dt.Columns("qty_on_hand").DataType.ToString.Trim)
            obj(7) = ReplaceNullValue(rw("qty_allocated").ToString.Trim, dt.Columns("qty_allocated").DataType.ToString.Trim)
            obj(8) = ReplaceNullValue(rw("qty_bkord").ToString.Trim, dt.Columns("qty_bkord").DataType.ToString.Trim)
            obj(9) = ReplaceNullValue(rw("qty_on_ord").ToString.Trim, dt.Columns("qty_on_ord").DataType.ToString.Trim)
            obj(10) = ReplaceNullValue(rw("reorder_lvl").ToString.Trim, dt.Columns("reorder_lvl").DataType.ToString.Trim)
            obj(11) = ReplaceNullValue(rw("ord_up_to_lvl").ToString.Trim, dt.Columns("ord_up_to_lvl").DataType.ToString.Trim)
            obj(12) = ReplaceNullValue(rw("price").ToString.Trim, dt.Columns("price").DataType.ToString.Trim)
            obj(13) = ReplaceNullValue(rw("avg_cost").ToString.Trim, dt.Columns("avg_cost").DataType.ToString.Trim)
            obj(14) = ReplaceNullValue(rw("last_cost").ToString.Trim, dt.Columns("last_cost").DataType.ToString.Trim)
            obj(15) = ReplaceNullValue(rw("std_cost").ToString.Trim, dt.Columns("std_cost").DataType.ToString.Trim)
            obj(16) = ReplaceNullValue(rw("prcs_apply_fg").ToString.Trim, dt.Columns("prcs_apply_fg").DataType.ToString.Trim)
            obj(17) = ReplaceNullValue(rw("discs_apply_fg").ToString.Trim, dt.Columns("discs_apply_fg").DataType.ToString.Trim)
            obj(18) = ReplaceNullValue(rw("starting_sls_dt").ToString.Trim, dt.Columns("starting_sls_dt").DataType.ToString.Trim)
            obj(19) = ReplaceNullValue(rw("ending_sls_dt").ToString.Trim, dt.Columns("ending_sls_dt").DataType.ToString.Trim)
            obj(20) = ReplaceNullValue(rw("last_sold_dt").ToString.Trim, dt.Columns("last_sold_dt").DataType.ToString.Trim)
            obj(21) = ReplaceNullValue(rw("sls_price").ToString.Trim, dt.Columns("sls_price").DataType.ToString.Trim)
            obj(22) = ReplaceNullValue(rw("qty_last_sold").ToString.Trim, dt.Columns("qty_last_sold").DataType.ToString.Trim)
            obj(23) = ReplaceNullValue(rw("cycle_count_cd").ToString.Trim, dt.Columns("cycle_count_cd").DataType.ToString.Trim)
            obj(24) = ReplaceNullValue(rw("last_count_dt").ToString.Trim, dt.Columns("last_count_dt").DataType.ToString.Trim)
            obj(25) = ReplaceNullValue(rw("tms_cntd_ytd").ToString.Trim, dt.Columns("tms_cntd_ytd").DataType.ToString.Trim)
            obj(26) = ReplaceNullValue(rw("pct_err_last_cnt").ToString.Trim, dt.Columns("pct_err_last_cnt").DataType.ToString.Trim)
            obj(27) = ReplaceNullValue(rw("frz_cost").ToString.Trim, dt.Columns("frz_cost").DataType.ToString.Trim)
            obj(28) = ReplaceNullValue(rw("frz_qty").ToString.Trim, dt.Columns("frz_qty").DataType.ToString.Trim)
            obj(29) = ReplaceNullValue(rw("frz_dt").ToString.Trim, dt.Columns("frz_dt").DataType.ToString.Trim)
            obj(30) = ReplaceNullValue(rw("frz_tm").ToString.Trim, dt.Columns("frz_tm").DataType.ToString.Trim)
            obj(31) = ReplaceNullValue(rw("usage_ptd").ToString.Trim, dt.Columns("usage_ptd").DataType.ToString.Trim)
            obj(32) = ReplaceNullValue(rw("qty_sld_ptd").ToString.Trim, dt.Columns("qty_sld_ptd").DataType.ToString.Trim)
            obj(33) = ReplaceNullValue(rw("qty_scrp_ptd").ToString.Trim, dt.Columns("qty_scrp_ptd").DataType.ToString.Trim)
            obj(34) = ReplaceNullValue(rw("sls_ptd").ToString.Trim, dt.Columns("sls_ptd").DataType.ToString.Trim)
            obj(35) = ReplaceNullValue(rw("cost_ptd").ToString.Trim, dt.Columns("cost_ptd").DataType.ToString.Trim)
            obj(36) = ReplaceNullValue(rw("usage_ytd").ToString.Trim, dt.Columns("usage_ytd").DataType.ToString.Trim)
            obj(37) = ReplaceNullValue(rw("qty_sold_ytd").ToString.Trim, dt.Columns("qty_sold_ytd").DataType.ToString.Trim)
            obj(38) = ReplaceNullValue(rw("qty_scrp_ytd").ToString.Trim, dt.Columns("qty_scrp_ytd").DataType.ToString.Trim)
            obj(39) = ReplaceNullValue(rw("qty_returned_ytd").ToString.Trim, dt.Columns("qty_returned_ytd").DataType.ToString.Trim)
            obj(40) = ReplaceNullValue(rw("sls_ytd").ToString.Trim, dt.Columns("sls_ytd").DataType.ToString.Trim)
            obj(41) = ReplaceNullValue(rw("cost_ytd").ToString.Trim, dt.Columns("cost_ytd").DataType.ToString.Trim)
            obj(42) = ReplaceNullValue(rw("prior_year_usage").ToString.Trim, dt.Columns("prior_year_usage").DataType.ToString.Trim)
            obj(43) = ReplaceNullValue(rw("qty_sold_last_yr").ToString.Trim, dt.Columns("qty_sold_last_yr").DataType.ToString.Trim)
            obj(44) = ReplaceNullValue(rw("qty_scrp_last_yr").ToString.Trim, dt.Columns("qty_scrp_last_yr").DataType.ToString.Trim)
            obj(45) = ReplaceNullValue(rw("prior_year_sls").ToString.Trim, dt.Columns("prior_year_sls").DataType.ToString.Trim)
            obj(46) = ReplaceNullValue(rw("cost_last_yr").ToString.Trim, dt.Columns("cost_last_yr").DataType.ToString.Trim)
            obj(47) = ReplaceNullValue(rw("recom_min_ord").ToString.Trim, dt.Columns("recom_min_ord").DataType.ToString.Trim)
            obj(48) = ReplaceNullValue(rw("economic_ord_qty").ToString.Trim, dt.Columns("economic_ord_qty").DataType.ToString.Trim)
            obj(49) = ReplaceNullValue(rw("avg_usage").ToString.Trim, dt.Columns("avg_usage").DataType.ToString.Trim)
            obj(50) = ReplaceNullValue(rw("po_lead_tm").ToString.Trim, dt.Columns("po_lead_tm").DataType.ToString.Trim)
            obj(51) = ReplaceNullValue(rw("byr_plnr").ToString.Trim, dt.Columns("byr_plnr").DataType.ToString.Trim)
            obj(52) = ReplaceNullValue(rw("doc_to_stk_ld_tm").ToString.Trim, dt.Columns("doc_to_stk_ld_tm").DataType.ToString.Trim)
            obj(53) = ReplaceNullValue(rw("rollup_prc").ToString.Trim, dt.Columns("rollup_prc").DataType.ToString.Trim)
            obj(54) = ReplaceNullValue(rw("target_margin").ToString.Trim, dt.Columns("target_margin").DataType.ToString.Trim)
            obj(55) = ReplaceNullValue(rw("inv_class").ToString.Trim, dt.Columns("inv_class").DataType.ToString.Trim)
            obj(56) = ReplaceNullValue(rw("po_min").ToString.Trim, dt.Columns("po_min").DataType.ToString.Trim)
            obj(57) = ReplaceNullValue(rw("po_max").ToString.Trim, dt.Columns("po_max").DataType.ToString.Trim)
            obj(58) = ReplaceNullValue(rw("safety_stk").ToString.Trim, dt.Columns("safety_stk").DataType.ToString.Trim)
            obj(59) = ReplaceNullValue(rw("avg_frcst_error").ToString.Trim, dt.Columns("avg_frcst_error").DataType.ToString.Trim)
            obj(60) = ReplaceNullValue(rw("sum_of_errors").ToString.Trim, dt.Columns("sum_of_errors").DataType.ToString.Trim)
            obj(61) = ReplaceNullValue(rw("usg_wght_fctr").ToString.Trim, dt.Columns("usg_wght_fctr").DataType.ToString.Trim)
            obj(62) = ReplaceNullValue(rw("safety_fctr").ToString.Trim, dt.Columns("safety_fctr").DataType.ToString.Trim)
            obj(63) = ReplaceNullValue(rw("usage_filter").ToString.Trim, dt.Columns("usage_filter").DataType.ToString.Trim)
            obj(64) = ReplaceNullValue(rw("po_mult").ToString.Trim, dt.Columns("po_mult").DataType.ToString.Trim)
            obj(65) = ReplaceNullValue(rw("active_ords").ToString.Trim, dt.Columns("active_ords").DataType.ToString.Trim)
            obj(66) = ReplaceNullValue(rw("vend_no").ToString.Trim, dt.Columns("vend_no").DataType.ToString.Trim)
            obj(67) = ReplaceNullValue(rw("tax_sched").ToString.Trim, dt.Columns("tax_sched").DataType.ToString.Trim)
            obj(68) = ReplaceNullValue(rw("prod_cat").ToString.Trim, dt.Columns("prod_cat").DataType.ToString.Trim)
            obj(69) = ReplaceNullValue(rw("picking_seq").ToString.Trim, dt.Columns("picking_seq").DataType.ToString.Trim)
            obj(70) = ReplaceNullValue(rw("cube_width_uom").ToString.Trim, dt.Columns("cube_width_uom").DataType.ToString.Trim)
            obj(71) = ReplaceNullValue(rw("cube_length_uom").ToString.Trim, dt.Columns("cube_length_uom").DataType.ToString.Trim)
            obj(72) = ReplaceNullValue(rw("cube_height_uom").ToString.Trim, dt.Columns("cube_height_uom").DataType.ToString.Trim)
            obj(73) = ReplaceNullValue(rw("cube_width").ToString.Trim, dt.Columns("cube_width").DataType.ToString.Trim)
            obj(74) = ReplaceNullValue(rw("cube_length").ToString.Trim, dt.Columns("cube_length").DataType.ToString.Trim)
            obj(75) = ReplaceNullValue(rw("cube_height").ToString.Trim, dt.Columns("cube_height").DataType.ToString.Trim)
            obj(76) = ReplaceNullValue(rw("cube_qty_per").ToString.Trim, dt.Columns("cube_qty_per").DataType.ToString.Trim)
            obj(77) = ReplaceNullValue(rw("user_def_fld_1").ToString.Trim, dt.Columns("user_def_fld_1").DataType.ToString.Trim)
            obj(78) = ReplaceNullValue(rw("user_def_fld_2").ToString.Trim, dt.Columns("user_def_fld_2").DataType.ToString.Trim)
            obj(79) = ReplaceNullValue(rw("user_def_fld_3").ToString.Trim, dt.Columns("user_def_fld_3").DataType.ToString.Trim)
            obj(80) = ReplaceNullValue(rw("user_def_fld_4").ToString.Trim, dt.Columns("user_def_fld_4").DataType.ToString.Trim)
            obj(81) = ReplaceNullValue(rw("user_def_fld_5").ToString.Trim, dt.Columns("user_def_fld_5").DataType.ToString.Trim)
            obj(82) = ReplaceNullValue(rw("landed_cost_cd").ToString.Trim, dt.Columns("landed_cost_cd").DataType.ToString.Trim)
            obj(83) = ReplaceNullValue(rw("landed_cost_cd_2").ToString.Trim, dt.Columns("landed_cost_cd_2").DataType.ToString.Trim)
            obj(84) = ReplaceNullValue(rw("landed_cost_cd_3").ToString.Trim, dt.Columns("landed_cost_cd_3").DataType.ToString.Trim)
            obj(85) = ReplaceNullValue(rw("landed_cost_cd_4").ToString.Trim, dt.Columns("landed_cost_cd_4").DataType.ToString.Trim)
            obj(86) = ReplaceNullValue(rw("landed_cost_cd_5").ToString.Trim, dt.Columns("landed_cost_cd_5").DataType.ToString.Trim)
            obj(87) = ReplaceNullValue(rw("landed_cost_cd_6").ToString.Trim, dt.Columns("landed_cost_cd_6").DataType.ToString.Trim)
            obj(88) = ReplaceNullValue(rw("landed_cost_cd_7").ToString.Trim, dt.Columns("landed_cost_cd_7").DataType.ToString.Trim)
            obj(89) = ReplaceNullValue(rw("landed_cost_cd_8").ToString.Trim, dt.Columns("landed_cost_cd_8").DataType.ToString.Trim)
            obj(90) = ReplaceNullValue(rw("landed_cost_cd_9").ToString.Trim, dt.Columns("landed_cost_cd_9").DataType.ToString.Trim)
            obj(91) = ReplaceNullValue(rw("landed_cost_cd_10").ToString.Trim, dt.Columns("landed_cost_cd_10").DataType.ToString.Trim)
            obj(92) = ReplaceNullValue(rw("loc_qty_fld").ToString.Trim, dt.Columns("loc_qty_fld").DataType.ToString.Trim)
            obj(93) = ReplaceNullValue(rw("tag_qty").ToString.Trim, dt.Columns("tag_qty").DataType.ToString.Trim)
            obj(94) = ReplaceNullValue(rw("tag_cost").ToString.Trim, dt.Columns("tag_cost").DataType.ToString.Trim)
            obj(95) = ReplaceNullValue(rw("tag_frz_dt").ToString.Trim, dt.Columns("tag_frz_dt").DataType.ToString.Trim)
            obj(96) = ReplaceNullValue(rw("filler_0002").ToString.Trim, dt.Columns("filler_0002").DataType.ToString.Trim)


        Next

        invloc = NewInventoryLocation(
        obj(0).ToString,
        obj(1).ToString,
        obj(2).ToString,
        obj(3).ToString,
        obj(4).ToString,
        obj(5).ToString,
        Convert.ToDecimal(obj(6)).ToString.Trim,
        Convert.ToDecimal(obj(7)).ToString.Trim,
        Convert.ToDecimal(obj(8)).ToString.Trim,
        Convert.ToDecimal(obj(9)).ToString.Trim,
        Convert.ToDecimal(obj(10)).ToString.Trim,
        Convert.ToDecimal(obj(11)).ToString.Trim,
        Convert.ToDecimal(obj(12)).ToString.Trim,
        Convert.ToDecimal(obj(13)).ToString.Trim,
        Convert.ToDecimal(obj(14)).ToString.Trim,
        Convert.ToDecimal(obj(15)).ToString.Trim,
        obj(16).ToString,
        obj(17).ToString,
        Convert.ToInt32(obj(18)).ToString.Trim,
        Convert.ToInt32(obj(19)).ToString.Trim,
        Convert.ToInt32(obj(20)).ToString.Trim,
        Convert.ToDecimal(obj(21)).ToString.Trim,
        Convert.ToDecimal(obj(22)).ToString.Trim,
        obj(23).ToString,
        Convert.ToInt32(obj(24)).ToString.Trim,
        Convert.ToInt32(obj(25)).ToString.Trim,
        Convert.ToDecimal(obj(26)).ToString.Trim,
        Convert.ToDecimal(obj(27)).ToString.Trim,
        Convert.ToDecimal(obj(28)).ToString.Trim,
        Convert.ToInt32(obj(29)).ToString.Trim,
        Convert.ToInt32(obj(30)).ToString.Trim,
        Convert.ToDecimal(obj(31)).ToString.Trim,
        Convert.ToDecimal(obj(32)).ToString.Trim,
        Convert.ToDecimal(obj(33)).ToString.Trim,
        Convert.ToDecimal(obj(34)).ToString.Trim,
        Convert.ToDecimal(obj(35)).ToString.Trim,
        Convert.ToDecimal(obj(36)).ToString.Trim,
        Convert.ToDecimal(obj(37)).ToString.Trim,
        Convert.ToDecimal(obj(38)).ToString.Trim,
        Convert.ToDecimal(obj(39)).ToString.Trim,
        Convert.ToDecimal(obj(40)).ToString.Trim,
        Convert.ToDecimal(obj(41)).ToString.Trim,
        Convert.ToDecimal(obj(42)).ToString.Trim,
        Convert.ToDecimal(obj(43)).ToString.Trim,
        Convert.ToDecimal(obj(44)).ToString.Trim,
        Convert.ToDecimal(obj(45)).ToString.Trim,
        Convert.ToDecimal(obj(46)).ToString.Trim,
        Convert.ToDecimal(obj(47)).ToString.Trim,
        Convert.ToDecimal(obj(48)).ToString.Trim,
        Convert.ToDecimal(obj(49)).ToString.Trim,
        Convert.ToInt32(obj(50)).ToString.Trim,
        obj(51).ToString,
        Convert.ToInt32(obj(52)).ToString.Trim,
        obj(53).ToString,
        Convert.ToInt32(obj(54)).ToString.Trim,
        obj(55).ToString,
        Convert.ToDecimal(obj(56)).ToString.Trim,
        Convert.ToDecimal(obj(57)).ToString.Trim,
        Convert.ToDecimal(obj(58)).ToString.Trim,
        Convert.ToDecimal(obj(59)).ToString.Trim,
        Convert.ToDecimal(obj(60)).ToString.Trim,
        Convert.ToDecimal(obj(61)).ToString.Trim,
        Convert.ToDecimal(obj(62)).ToString.Trim,
        Convert.ToDecimal(obj(63)).ToString.Trim,
        Convert.ToInt32(obj(64)).ToString.Trim,
        Convert.ToInt32(obj(65)).ToString.Trim,
        obj(66).ToString,
        obj(67).ToString,
        obj(68).ToString,
        obj(69).ToString,
        obj(70).ToString,
        obj(71).ToString,
        obj(72).ToString,
        Convert.ToDecimal(obj(73)).ToString.Trim,
        Convert.ToDecimal(obj(74)).ToString.Trim,
        Convert.ToDecimal(obj(75)).ToString.Trim,
        Convert.ToDecimal(obj(76)).ToString.Trim,
        obj(77).ToString,
        obj(78).ToString,
        obj(79).ToString,
        obj(80).ToString,
        obj(81).ToString,
        obj(82).ToString,
        obj(83).ToString,
        obj(84).ToString,
        obj(85).ToString,
        obj(86).ToString,
        obj(87).ToString,
        obj(88).ToString,
        obj(89).ToString,
        obj(90).ToString,
        obj(91).ToString,
        Convert.ToDecimal(obj(92)).ToString.Trim,
        Convert.ToDecimal(obj(93)).ToString.Trim,
        Convert.ToDecimal(obj(94)).ToString.Trim,
        Convert.ToInt32(obj(95)).ToString.Trim,
        obj(96).ToString
)

        Return invloc

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
                    If TypeOf (value) Is String And value = "" Then
                        val = 0
                    Else
                        val = value
                    End If

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
