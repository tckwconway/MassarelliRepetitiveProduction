Imports System.ComponentModel
Imports System.Data.DataRowView
Imports System.Data.SqlClient
Imports System.Data
Imports System.Text
Imports System.Xml
Imports System.IO
Imports System.Data.SqlTypes

Public Class IMControlSettings

    Public im_ctl_key_1 As Integer
    Public ave_last_cst_fg As String
    Public multi_acct_fg As String
    Public loc As String
    Public mat_cost_type As String
    Public audit_trail_fg As String
    Public no_of_days_in_prd As Decimal
    Public curr_prd As Integer
    Public non_stk_trx_aud As String
    Public chg_fg As String
    Public no_of_prds As Integer
    Public trx_aud_fg As String
    Public next_doc_no As Integer
    Public aud_file_bb_dt As Integer
    Public user_name As String
    Public online_upd_fg As String
    Public update_dist_phy_cnt As String
    Public proces_nonstk_bomp As String
    Public distribute_qty_amt As String
    Public use_job_nos_fg As String
    Public next_tag_no As Integer
    Public next_text_no As Integer
    Public mn_no As String
    Public sb_no As String
    Public dp_no As String
    Public receiving_mn_no As String
    Public receiving_sb_no As String
    Public receiving_dp_no As String
    Public issue_mn_no As String
    Public issue_sb_no As String
    Public issue_dp_no As String
    Public receipt_mn_no As String
    Public receipt_sb_no As String
    Public receipt_dp_no As String
    Public qty_adj_mn_no As String
    Public qty_adj_sb_no As String
    Public qty_adj_dp_no As String
    Public cost_adj_mn_no As String
    Public cost_adj_sb_no As String
    Public cost_adj_dp_no As String
    Public wip_var_mn_no As String
    Public wip_var_sb_no As String
    Public wip_var_dp_no As String
    Public ppv_var_mn_no As String
    Public ppv_var_sb_no As String
    Public ppv_var_dp_no As String
    Public ppv_qty_var_mn_no As String
    Public ppv_qty_var_sb_no As String
    Public ppv_qty_var_dp_no As String
    Public xfr_var_mn_no As String
    Public xfr_var_sb_no As String
    Public xfr_var_dp_no As String
    Public cyc_phy_cnt_mn_no As String
    Public cyc_phy_cnt_sb_no As String
    Public cyc_phy_cnt_dp_no As String
    Public use_mult_bin_fg As String
    Public bin_meth As String
    Public use_ri_mrb_loc As String
    Public use_land_cst_fg As String
    Public use_ser_lot_fg As String
    Public ser_lot_cnt_tags As String
    Public use_disc_cst_fg As String
    Public ri_accr_mn_no As String
    Public ri_accr_sb_no As String
    Public ri_accr_dp_no As String
    Public ser_no_auto_assign As String
    Public ser_no_static_len As Integer
    Public ser_no_static_val As String
    Public ser_no_numeric_len As Integer
    Public ser_no_numeric_val As Decimal
    Public filler_0001 As String
    Public lot_no_auto_assign As String
    Public lot_no_static_len As Integer
    Public lot_no_static_val As String
    Public lot_no_numeric_len As Integer
    Public lot_no_numeric_val As Decimal
    Public filler_0002 As String
    Public disab_prot_fld_fg As String
    Public batch_ctl_fg As String
    Public alloc_meth_fg As String
    Public dflt_tag_frm As Integer
    Public dflt_lbl_frm As Integer
    Public prd_size As Integer
    Public day_of_week As Integer
    Public bin_receipt_meth As String
    Public ser_lot_rec_meth As String
    Public lf_init_fg As String
    Public note_lit_1 As String
    Public note_lit_2 As String
    Public note_lit_3 As String
    Public note_lit_4 As String
    Public note_lit_5 As String
    Public dt_lit As String
    Public amt_lit As String
    Public use_expire_sl_fg As String
    Public show_zero_qty_bins As String
    Public auto_except_fg As String
    Public insuff_display_fg As String
    Public bypass_gl_dist_fg As String
    Public atp_def_atp_prds As Integer
    Public atp_def_use_stk As String
    Public atp_def_use_ord As String
    Public atp_def_use_coq As String
    Public atp_def_min_cd As Integer
    Public atp_def_use_cob As String
    Public atp_def_use_unr As String
    Public atp_def_use_fp As String
    Public atp_def_use_cp As String
    Public ser_lot_range_fg As String
    Public filler_0003 As String

    Const TABLE As String = "IMCTLFIL_SQL"

    Private Function NewIMControlSettings(ByVal m_im_ctl_key_1 As Integer, ByVal m_ave_last_cst_fg As String, ByVal m_multi_acct_fg As String, ByVal m_loc As String, ByVal m_mat_cost_type As String, _
                                            ByVal m_audit_trail_fg As String, ByVal m_no_of_days_in_prd As Decimal, ByVal m_curr_prd As Integer, ByVal m_non_stk_trx_aud As String, ByVal m_chg_fg As String, _
                                            ByVal m_no_of_prds As Integer, ByVal m_trx_aud_fg As String, ByVal m_next_doc_no As Integer, ByVal m_aud_file_bb_dt As Integer, ByVal m_user_name As String, _
                                            ByVal m_online_upd_fg As String, ByVal m_update_dist_phy_cnt As String, ByVal m_proces_nonstk_bomp As String, ByVal m_distribute_qty_amt As String, ByVal m_use_job_nos_fg As String, _
                                            ByVal m_next_aag_no As Integer, ByVal m_next_text_no As Integer, ByVal m_mn_no As String, ByVal m_sb_no As String, ByVal m_dp_no As String, ByVal m_receiving_mn_no As String, _
                                            ByVal m_receiving_sb_no As String, ByVal m_receiving_dp_no As String, ByVal m_issue_mn_no As String, ByVal m_issue_sb_no As String, ByVal m_issue_dp_no As String, ByVal m_receipt_mn_no As String, _
                                            ByVal m_receipt_sb_no As String, ByVal m_receipt_dp_no As String, ByVal m_qty_adj_mn_no As String, ByVal m_qty_adj_sb_no As String, ByVal m_qty_adj_dp_no As String, ByVal m_cost_adj_mn_no As String, _
                                            ByVal m_cost_adj_sb_no As String, ByVal m_cost_adj_dp_no As String, ByVal m_wip_var_mn_no As String, ByVal m_wip_var_sb_no As String, ByVal m_wip_var_dp_no As String, _
                                            ByVal m_ppv_var_mn_no As String, ByVal m_ppv_var_sb_no As String, ByVal m_ppv_var_dp_no As String, ByVal m_ppv_qty_var_mn_no As String, ByVal m_ppv_qty_var_sb_no As String, _
                                            ByVal m_ppv_qty_var_dp_no As String, ByVal m_xfr_var_mn_no As String, ByVal m_xfr_var_sb_no As String, ByVal m_xfr_var_dp_no As String, ByVal m_cyc_phy_cnt_mn_no As String, _
                                            ByVal m_cyc_phy_cnt_sb_no As String, ByVal m_cyc_phy_cnt_dp_no As String, ByVal m_use_mult_bin_fg As String, ByVal m_bin_meth As String, ByVal m_use_ri_mrb_loc As String, _
                                            ByVal m_use_land_cst_fg As String, ByVal m_use_ser_lot_fg As String, ByVal m_ser_lot_cnt_tags As String, ByVal m_use_disc_cst_fg As String, ByVal m_ri_accr_mn_no As String, _
                                            ByVal m_ri_accr_sb_no As String, ByVal m_ri_accr_dp_no As String, ByVal m_ser_no_auto_assign As String, ByVal m_ser_no_static_len As Integer, ByVal m_ser_no_static_val As String, _
                                            ByVal m_ser_no_numeric_len As Integer, ByVal m_ser_no_numeric_val As Decimal, ByVal m_filler_0001 As String, ByVal m_lot_no_auto_assign As String, ByVal m_lot_no_static_len As Integer, _
                                            ByVal m_lot_no_static_val As String, ByVal m_lot_no_numeric_len As Integer, ByVal m_lot_no_numeric_val As Decimal, ByVal m_filler_0002 As String, ByVal m_disab_prot_fld_fg As String, _
                                            ByVal m_batch_ctl_fg As String, ByVal m_alloc_meth_fg As String, ByVal m_dflt_tag_frm As Integer, ByVal m_dflt_lbl_frm As Integer, ByVal m_prd_size As Integer, ByVal m_day_of_week As Integer, _
                                            ByVal m_bin_receipt_meth As String, ByVal m_ser_lot_rec_meth As String, ByVal m_lf_init_fg As String, ByVal m_note_lit_1 As String, ByVal m_note_lit_2 As String, ByVal m_note_lit_3 As String, _
                                            ByVal m_note_lit_4 As String, ByVal m_note_lit_5 As String, ByVal m_dt_lit As String, ByVal m_amt_lit As String, ByVal m_use_expire_sl_fg As String, ByVal m_show_zero_qty_bins As String, _
                                            ByVal m_auto_except_fg As String, ByVal m_insuff_display_fg As String, ByVal m_bypass_gl_dist_fg As String, ByVal m_atp_def_atp_prds As Integer, ByVal m_atp_def_use_stk As String, ByVal m_atp_def_use_ord As String, _
                                            ByVal m_atp_def_use_coq As String, ByVal m_atp_def_min_cd As Integer, ByVal m_atp_def_use_cob As String, ByVal m_atp_def_use_unr As String, ByVal m_atp_def_use_fp As String, ByVal m_atp_def_use_cp As String, _
                                            ByVal m_ser_lot_range_fg As String, ByVal m_filler_0003 As String)



        Dim imctl As New IMControlSettings

        imctl.im_ctl_key_1 = m_im_ctl_key_1
        imctl.ave_last_cst_fg = m_ave_last_cst_fg
        imctl.multi_acct_fg = m_multi_acct_fg
        imctl.loc = m_loc
        imctl.mat_cost_type = m_mat_cost_type
        imctl.audit_trail_fg = m_audit_trail_fg
        imctl.no_of_days_in_prd = m_no_of_days_in_prd
        imctl.curr_prd = m_curr_prd
        imctl.non_stk_trx_aud = m_non_stk_trx_aud
        imctl.chg_fg = m_chg_fg
        imctl.no_of_prds = m_no_of_prds
        imctl.trx_aud_fg = m_trx_aud_fg
        imctl.next_doc_no = m_next_doc_no
        imctl.aud_file_bb_dt = m_aud_file_bb_dt
        imctl.user_name = m_user_name
        imctl.online_upd_fg = m_online_upd_fg
        imctl.update_dist_phy_cnt = m_update_dist_phy_cnt
        imctl.proces_nonstk_bomp = m_proces_nonstk_bomp
        imctl.distribute_qty_amt = m_distribute_qty_amt
        imctl.use_job_nos_fg = m_use_job_nos_fg
        imctl.next_tag_no = m_next_text_no
        imctl.next_text_no = m_next_text_no
        imctl.mn_no = m_mn_no
        imctl.sb_no = m_sb_no
        imctl.dp_no = m_dp_no
        imctl.receiving_mn_no = m_receiving_mn_no
        imctl.receiving_sb_no = m_receiving_sb_no
        imctl.receiving_dp_no = m_receiving_dp_no
        imctl.issue_mn_no = m_issue_mn_no
        imctl.issue_sb_no = m_issue_sb_no
        imctl.issue_dp_no = m_issue_dp_no
        imctl.receipt_mn_no = m_receipt_mn_no
        imctl.receipt_sb_no = m_receipt_sb_no
        imctl.receipt_dp_no = m_receipt_dp_no
        imctl.qty_adj_mn_no = m_qty_adj_mn_no
        imctl.qty_adj_sb_no = m_qty_adj_sb_no
        imctl.qty_adj_dp_no = m_qty_adj_dp_no
        imctl.cost_adj_mn_no = m_cost_adj_mn_no
        imctl.cost_adj_sb_no = m_cost_adj_sb_no
        imctl.cost_adj_dp_no = m_cost_adj_dp_no
        imctl.wip_var_mn_no = m_wip_var_mn_no
        imctl.wip_var_sb_no = m_wip_var_sb_no
        imctl.wip_var_dp_no = m_wip_var_dp_no
        imctl.ppv_var_mn_no = m_ppv_var_mn_no
        imctl.ppv_var_sb_no = m_ppv_var_sb_no
        imctl.ppv_var_dp_no = m_ppv_var_dp_no
        imctl.ppv_qty_var_mn_no = m_ppv_qty_var_mn_no
        imctl.ppv_qty_var_sb_no = m_ppv_qty_var_sb_no
        imctl.ppv_qty_var_dp_no = m_ppv_qty_var_dp_no
        imctl.xfr_var_mn_no = m_xfr_var_mn_no
        imctl.xfr_var_sb_no = m_xfr_var_sb_no
        imctl.xfr_var_dp_no = m_xfr_var_dp_no
        imctl.cyc_phy_cnt_mn_no = m_cyc_phy_cnt_mn_no
        imctl.cyc_phy_cnt_sb_no = m_cyc_phy_cnt_sb_no
        imctl.cyc_phy_cnt_dp_no = m_cyc_phy_cnt_dp_no
        imctl.use_mult_bin_fg = m_use_mult_bin_fg
        imctl.bin_meth = m_bin_meth
        imctl.use_ri_mrb_loc = m_use_ri_mrb_loc
        imctl.use_land_cst_fg = m_use_land_cst_fg
        imctl.use_ser_lot_fg = m_use_ser_lot_fg
        imctl.ser_lot_cnt_tags = m_ser_lot_cnt_tags
        imctl.use_disc_cst_fg = m_use_disc_cst_fg
        imctl.ri_accr_mn_no = m_ri_accr_mn_no
        imctl.ri_accr_sb_no = m_ri_accr_sb_no
        imctl.ri_accr_dp_no = m_ri_accr_dp_no
        imctl.ser_no_auto_assign = m_ser_no_auto_assign
        imctl.ser_no_static_len = m_ser_no_static_len
        imctl.ser_no_static_val = m_ser_no_static_val
        imctl.ser_no_numeric_len = m_ser_no_numeric_len
        imctl.ser_no_numeric_val = m_ser_no_numeric_val
        imctl.filler_0001 = m_filler_0001
        imctl.lot_no_auto_assign = m_lot_no_auto_assign
        imctl.lot_no_static_len = m_lot_no_static_len
        imctl.lot_no_static_val = m_lot_no_static_val
        imctl.lot_no_numeric_len = m_lot_no_numeric_len
        imctl.lot_no_numeric_val = m_lot_no_numeric_val
        imctl.filler_0002 = m_filler_0002
        imctl.disab_prot_fld_fg = m_disab_prot_fld_fg
        imctl.batch_ctl_fg = m_batch_ctl_fg
        imctl.alloc_meth_fg = m_alloc_meth_fg
        imctl.dflt_tag_frm = m_dflt_tag_frm
        imctl.dflt_lbl_frm = m_dflt_lbl_frm
        imctl.prd_size = m_prd_size
        imctl.day_of_week = m_day_of_week
        imctl.bin_receipt_meth = m_bin_receipt_meth
        imctl.ser_lot_rec_meth = m_ser_lot_rec_meth
        imctl.lf_init_fg = m_lf_init_fg
        imctl.note_lit_1 = m_note_lit_1
        imctl.note_lit_2 = m_note_lit_2
        imctl.note_lit_3 = m_note_lit_3
        imctl.note_lit_4 = m_note_lit_4
        imctl.note_lit_5 = m_note_lit_5
        imctl.dt_lit = m_dt_lit
        imctl.amt_lit = m_amt_lit
        imctl.use_expire_sl_fg = m_use_expire_sl_fg
        imctl.show_zero_qty_bins = m_show_zero_qty_bins
        imctl.auto_except_fg = m_auto_except_fg
        imctl.insuff_display_fg = m_insuff_display_fg
        imctl.bypass_gl_dist_fg = m_bypass_gl_dist_fg
        imctl.atp_def_atp_prds = m_atp_def_atp_prds
        imctl.atp_def_use_stk = m_atp_def_use_stk
        imctl.atp_def_use_ord = m_atp_def_use_ord
        imctl.atp_def_use_coq = m_atp_def_use_coq
        imctl.atp_def_min_cd = m_atp_def_min_cd
        imctl.atp_def_use_cob = m_atp_def_use_cob
        imctl.atp_def_use_unr = m_atp_def_use_unr
        imctl.atp_def_use_fp = m_atp_def_use_fp
        imctl.atp_def_use_cp = m_atp_def_use_cp
        imctl.ser_lot_range_fg = m_ser_lot_range_fg
        imctl.filler_0003 = m_filler_0003

        Return imctl

    End Function


    Public Function GetIMControlSetting(ByVal im_ctrl_key As String, ByVal cn As SqlConnection) As IMControlSettings

        'IM_CTL_KEY_1 for Massarelli is "1"

        Dim imctl As New IMControlSettings
        Dim dt As New DataTable
        Dim sSQL As String = "Select * from " & TABLE & " " _
                           & "Where IM_CTL_KEY_1 = '" & im_ctrl_key & "' "

        dt = DAC.ExecuteSQL_DataTable(sSQL, cn, TABLE)

        imctl = PopulateIMControlSettings(dt)

        Return imctl

    End Function


    Private Function PopulateIMControlSettings(ByVal dt As DataTable) As IMControlSettings
        Dim imctl As New IMControlSettings
        Dim obj(109) As Object
        For Each rw As DataRow In dt.Rows

            obj(0) = ReplaceNullValue(rw("im_ctl_key_1"), dt.Columns("im_ctl_key_1").DataType.ToString)
            obj(1) = ReplaceNullValue(rw("ave_last_cst_fg"), dt.Columns("ave_last_cst_fg").DataType.ToString)
            obj(2) = ReplaceNullValue(rw("multi_acct_fg"), dt.Columns("multi_acct_fg").DataType.ToString)
            obj(3) = ReplaceNullValue(rw("loc"), dt.Columns("loc").DataType.ToString)
            obj(4) = ReplaceNullValue(rw("mat_cost_type"), dt.Columns("mat_cost_type").DataType.ToString)
            obj(5) = ReplaceNullValue(rw("audit_trail_fg"), dt.Columns("audit_trail_fg").DataType.ToString)
            obj(6) = ReplaceNullValue(rw("no_of_days_in_prd"), dt.Columns("no_of_days_in_prd").DataType.ToString)
            obj(7) = ReplaceNullValue(rw("curr_prd"), dt.Columns("curr_prd").DataType.ToString)
            obj(8) = ReplaceNullValue(rw("non_stk_trx_aud"), dt.Columns("non_stk_trx_aud").DataType.ToString)
            obj(9) = ReplaceNullValue(rw("chg_fg"), dt.Columns("chg_fg").DataType.ToString)
            obj(10) = ReplaceNullValue(rw("no_of_prds"), dt.Columns("no_of_prds").DataType.ToString)
            obj(11) = ReplaceNullValue(rw("trx_aud_fg"), dt.Columns("trx_aud_fg").DataType.ToString)
            obj(12) = ReplaceNullValue(rw("next_doc_no"), dt.Columns("next_doc_no").DataType.ToString)
            obj(13) = ReplaceNullValue(rw("aud_file_bb_dt"), dt.Columns("aud_file_bb_dt").DataType.ToString)
            obj(14) = ReplaceNullValue(rw("user_name"), dt.Columns("user_name").DataType.ToString)
            obj(15) = ReplaceNullValue(rw("online_upd_fg"), dt.Columns("online_upd_fg").DataType.ToString)
            obj(16) = ReplaceNullValue(rw("update_dist_phy_cnt"), dt.Columns("update_dist_phy_cnt").DataType.ToString)
            obj(17) = ReplaceNullValue(rw("proces_nonstk_bomp"), dt.Columns("proces_nonstk_bomp").DataType.ToString)
            obj(18) = ReplaceNullValue(rw("distribute_qty_amt"), dt.Columns("distribute_qty_amt").DataType.ToString)
            obj(19) = ReplaceNullValue(rw("use_job_nos_fg"), dt.Columns("use_job_nos_fg").DataType.ToString)
            obj(20) = ReplaceNullValue(rw("next_tag_no"), dt.Columns("next_tag_no").DataType.ToString)
            obj(21) = ReplaceNullValue(rw("next_text_no"), dt.Columns("next_text_no").DataType.ToString)
            obj(22) = ReplaceNullValue(rw("mn_no"), dt.Columns("mn_no").DataType.ToString)
            obj(23) = ReplaceNullValue(rw("sb_no"), dt.Columns("sb_no").DataType.ToString)
            obj(24) = ReplaceNullValue(rw("dp_no"), dt.Columns("dp_no").DataType.ToString)
            obj(25) = ReplaceNullValue(rw("receiving_mn_no"), dt.Columns("receiving_mn_no").DataType.ToString)
            obj(26) = ReplaceNullValue(rw("receiving_sb_no"), dt.Columns("receiving_sb_no").DataType.ToString)
            obj(27) = ReplaceNullValue(rw("receiving_dp_no"), dt.Columns("receiving_dp_no").DataType.ToString)
            obj(28) = ReplaceNullValue(rw("issue_mn_no"), dt.Columns("issue_mn_no").DataType.ToString)
            obj(29) = ReplaceNullValue(rw("issue_sb_no"), dt.Columns("issue_sb_no").DataType.ToString)
            obj(30) = ReplaceNullValue(rw("issue_dp_no"), dt.Columns("issue_dp_no").DataType.ToString)
            obj(31) = ReplaceNullValue(rw("receipt_mn_no"), dt.Columns("receipt_mn_no").DataType.ToString)
            obj(32) = ReplaceNullValue(rw("receipt_sb_no"), dt.Columns("receipt_sb_no").DataType.ToString)
            obj(33) = ReplaceNullValue(rw("receipt_dp_no"), dt.Columns("receipt_dp_no").DataType.ToString)
            obj(34) = ReplaceNullValue(rw("qty_adj_mn_no"), dt.Columns("qty_adj_mn_no").DataType.ToString)
            obj(35) = ReplaceNullValue(rw("qty_adj_sb_no"), dt.Columns("qty_adj_sb_no").DataType.ToString)
            obj(36) = ReplaceNullValue(rw("qty_adj_dp_no"), dt.Columns("qty_adj_dp_no").DataType.ToString)
            obj(37) = ReplaceNullValue(rw("cost_adj_mn_no"), dt.Columns("cost_adj_mn_no").DataType.ToString)
            obj(38) = ReplaceNullValue(rw("cost_adj_sb_no"), dt.Columns("cost_adj_sb_no").DataType.ToString)
            obj(39) = ReplaceNullValue(rw("cost_adj_dp_no"), dt.Columns("cost_adj_dp_no").DataType.ToString)
            obj(40) = ReplaceNullValue(rw("wip_var_mn_no"), dt.Columns("wip_var_mn_no").DataType.ToString)
            obj(41) = ReplaceNullValue(rw("wip_var_sb_no"), dt.Columns("wip_var_sb_no").DataType.ToString)
            obj(42) = ReplaceNullValue(rw("wip_var_dp_no"), dt.Columns("wip_var_dp_no").DataType.ToString)
            obj(43) = ReplaceNullValue(rw("ppv_var_mn_no"), dt.Columns("ppv_var_mn_no").DataType.ToString)
            obj(44) = ReplaceNullValue(rw("ppv_var_sb_no"), dt.Columns("ppv_var_sb_no").DataType.ToString)
            obj(45) = ReplaceNullValue(rw("ppv_var_dp_no"), dt.Columns("ppv_var_dp_no").DataType.ToString)
            obj(46) = ReplaceNullValue(rw("ppv_qty_var_mn_no"), dt.Columns("ppv_qty_var_mn_no").DataType.ToString)
            obj(47) = ReplaceNullValue(rw("ppv_qty_var_sb_no"), dt.Columns("ppv_qty_var_sb_no").DataType.ToString)
            obj(48) = ReplaceNullValue(rw("ppv_qty_var_dp_no"), dt.Columns("ppv_qty_var_dp_no").DataType.ToString)
            obj(49) = ReplaceNullValue(rw("xfr_var_mn_no"), dt.Columns("xfr_var_mn_no").DataType.ToString)
            obj(50) = ReplaceNullValue(rw("xfr_var_sb_no"), dt.Columns("xfr_var_sb_no").DataType.ToString)
            obj(51) = ReplaceNullValue(rw("xfr_var_dp_no"), dt.Columns("xfr_var_dp_no").DataType.ToString)
            obj(52) = ReplaceNullValue(rw("cyc_phy_cnt_mn_no"), dt.Columns("cyc_phy_cnt_mn_no").DataType.ToString)
            obj(53) = ReplaceNullValue(rw("cyc_phy_cnt_sb_no"), dt.Columns("cyc_phy_cnt_sb_no").DataType.ToString)
            obj(54) = ReplaceNullValue(rw("cyc_phy_cnt_dp_no"), dt.Columns("cyc_phy_cnt_dp_no").DataType.ToString)
            obj(55) = ReplaceNullValue(rw("use_mult_bin_fg"), dt.Columns("use_mult_bin_fg").DataType.ToString)
            obj(56) = ReplaceNullValue(rw("bin_meth"), dt.Columns("bin_meth").DataType.ToString)
            obj(57) = ReplaceNullValue(rw("use_ri_mrb_loc"), dt.Columns("use_ri_mrb_loc").DataType.ToString)
            obj(58) = ReplaceNullValue(rw("use_land_cst_fg"), dt.Columns("use_land_cst_fg").DataType.ToString)
            obj(59) = ReplaceNullValue(rw("use_ser_lot_fg"), dt.Columns("use_ser_lot_fg").DataType.ToString)
            obj(60) = ReplaceNullValue(rw("ser_lot_cnt_tags"), dt.Columns("ser_lot_cnt_tags").DataType.ToString)
            obj(61) = ReplaceNullValue(rw("use_disc_cst_fg"), dt.Columns("use_disc_cst_fg").DataType.ToString)
            obj(62) = ReplaceNullValue(rw("ri_accr_mn_no"), dt.Columns("ri_accr_mn_no").DataType.ToString)
            obj(63) = ReplaceNullValue(rw("ri_accr_sb_no"), dt.Columns("ri_accr_sb_no").DataType.ToString)
            obj(64) = ReplaceNullValue(rw("ri_accr_dp_no"), dt.Columns("ri_accr_dp_no").DataType.ToString)
            obj(65) = ReplaceNullValue(rw("ser_no_auto_assign"), dt.Columns("ser_no_auto_assign").DataType.ToString)
            obj(66) = ReplaceNullValue(rw("ser_no_static_len"), dt.Columns("ser_no_static_len").DataType.ToString)
            obj(67) = ReplaceNullValue(rw("ser_no_static_val"), dt.Columns("ser_no_static_val").DataType.ToString)
            obj(68) = ReplaceNullValue(rw("ser_no_numeric_len"), dt.Columns("ser_no_numeric_len").DataType.ToString)
            obj(69) = ReplaceNullValue(rw("ser_no_numeric_val"), dt.Columns("ser_no_numeric_val").DataType.ToString)
            obj(70) = ReplaceNullValue(rw("filler_0001"), dt.Columns("filler_0001").DataType.ToString)
            obj(71) = ReplaceNullValue(rw("lot_no_auto_assign"), dt.Columns("lot_no_auto_assign").DataType.ToString)
            obj(72) = ReplaceNullValue(rw("lot_no_static_len"), dt.Columns("lot_no_static_len").DataType.ToString)
            obj(73) = ReplaceNullValue(rw("lot_no_static_val"), dt.Columns("lot_no_static_val").DataType.ToString)
            obj(74) = ReplaceNullValue(rw("lot_no_numeric_len"), dt.Columns("lot_no_numeric_len").DataType.ToString)
            obj(75) = ReplaceNullValue(rw("lot_no_numeric_val"), dt.Columns("lot_no_numeric_val").DataType.ToString)
            obj(76) = ReplaceNullValue(rw("filler_0002"), dt.Columns("filler_0002").DataType.ToString)
            obj(77) = ReplaceNullValue(rw("disab_prot_fld_fg"), dt.Columns("disab_prot_fld_fg").DataType.ToString)
            obj(78) = ReplaceNullValue(rw("batch_ctl_fg"), dt.Columns("batch_ctl_fg").DataType.ToString)
            obj(79) = ReplaceNullValue(rw("alloc_meth_fg"), dt.Columns("alloc_meth_fg").DataType.ToString)
            obj(80) = ReplaceNullValue(rw("dflt_tag_frm"), dt.Columns("dflt_tag_frm").DataType.ToString)
            obj(81) = ReplaceNullValue(rw("dflt_lbl_frm"), dt.Columns("dflt_lbl_frm").DataType.ToString)
            obj(82) = ReplaceNullValue(rw("prd_size"), dt.Columns("prd_size").DataType.ToString)
            obj(83) = ReplaceNullValue(rw("day_of_week"), dt.Columns("day_of_week").DataType.ToString)
            obj(84) = ReplaceNullValue(rw("bin_receipt_meth"), dt.Columns("bin_receipt_meth").DataType.ToString)
            obj(85) = ReplaceNullValue(rw("ser_lot_rec_meth"), dt.Columns("ser_lot_rec_meth").DataType.ToString)
            obj(86) = ReplaceNullValue(rw("lf_init_fg"), dt.Columns("lf_init_fg").DataType.ToString)
            obj(87) = ReplaceNullValue(rw("note_lit_1"), dt.Columns("note_lit_1").DataType.ToString)
            obj(88) = ReplaceNullValue(rw("note_lit_2"), dt.Columns("note_lit_2").DataType.ToString)
            obj(89) = ReplaceNullValue(rw("note_lit_3"), dt.Columns("note_lit_3").DataType.ToString)
            obj(90) = ReplaceNullValue(rw("note_lit_4"), dt.Columns("note_lit_4").DataType.ToString)
            obj(91) = ReplaceNullValue(rw("note_lit_5"), dt.Columns("note_lit_5").DataType.ToString)
            obj(92) = ReplaceNullValue(rw("dt_lit"), dt.Columns("dt_lit").DataType.ToString)
            obj(93) = ReplaceNullValue(rw("amt_lit"), dt.Columns("amt_lit").DataType.ToString)
            obj(94) = ReplaceNullValue(rw("use_expire_sl_fg"), dt.Columns("use_expire_sl_fg").DataType.ToString)
            obj(95) = ReplaceNullValue(rw("show_zero_qty_bins"), dt.Columns("show_zero_qty_bins").DataType.ToString)
            obj(96) = ReplaceNullValue(rw("auto_except_fg"), dt.Columns("auto_except_fg").DataType.ToString)
            obj(97) = ReplaceNullValue(rw("insuff_display_fg"), dt.Columns("insuff_display_fg").DataType.ToString)
            obj(98) = ReplaceNullValue(rw("bypass_gl_dist_fg"), dt.Columns("bypass_gl_dist_fg").DataType.ToString)
            obj(99) = ReplaceNullValue(rw("atp_def_atp_prds"), dt.Columns("atp_def_atp_prds").DataType.ToString)
            obj(100) = ReplaceNullValue(rw("atp_def_use_stk"), dt.Columns("atp_def_use_stk").DataType.ToString)
            obj(101) = ReplaceNullValue(rw("atp_def_use_ord"), dt.Columns("atp_def_use_ord").DataType.ToString)
            obj(102) = ReplaceNullValue(rw("atp_def_use_coq"), dt.Columns("atp_def_use_coq").DataType.ToString)
            obj(103) = ReplaceNullValue(rw("atp_def_min_cd"), dt.Columns("atp_def_min_cd").DataType.ToString)
            obj(104) = ReplaceNullValue(rw("atp_def_use_cob"), dt.Columns("atp_def_use_cob").DataType.ToString)
            obj(105) = ReplaceNullValue(rw("atp_def_use_unr"), dt.Columns("atp_def_use_unr").DataType.ToString)
            obj(106) = ReplaceNullValue(rw("atp_def_use_fp"), dt.Columns("atp_def_use_fp").DataType.ToString)
            obj(107) = ReplaceNullValue(rw("atp_def_use_cp"), dt.Columns("atp_def_use_cp").DataType.ToString)
            obj(108) = ReplaceNullValue(rw("ser_lot_range_fg"), dt.Columns("ser_lot_range_fg").DataType.ToString)
            obj(109) = ReplaceNullValue(rw("filler_0003"), dt.Columns("filler_0003").DataType.ToString)

            imctl = NewIMControlSettings(
            Convert.ToInt32(obj(0)),
            obj(1).ToString,
            obj(2).ToString,
            obj(3).ToString,
            obj(4).ToString,
            obj(5).ToString,
            Convert.ToDecimal(obj(6)),
            Convert.ToInt32(obj(7)),
            obj(8).ToString,
            obj(9).ToString,
            Convert.ToInt32(obj(10)),
            obj(11).ToString,
            Convert.ToInt32(obj(12)),
            Convert.ToInt32(obj(13)),
            obj(14).ToString,
            obj(15).ToString,
            obj(16).ToString,
            obj(17).ToString,
            obj(18).ToString,
            obj(19).ToString,
            Convert.ToInt32(obj(20)),
            Convert.ToInt32(obj(21)),
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
            obj(52).ToString,
            obj(53).ToString,
            obj(54).ToString,
            obj(55).ToString,
            obj(56).ToString,
            obj(57).ToString,
            obj(58).ToString,
            obj(59).ToString,
            obj(60).ToString,
            obj(61).ToString,
            obj(62).ToString,
            obj(63).ToString,
            obj(64).ToString,
            obj(65).ToString,
            Convert.ToInt32(obj(66)),
            obj(67).ToString,
            Convert.ToInt32(obj(68)),
            Convert.ToDecimal(obj(69)),
            obj(70).ToString,
            obj(71).ToString,
            Convert.ToInt32(obj(72)),
            obj(73).ToString,
            Convert.ToInt32(obj(74)),
            Convert.ToDecimal(obj(75)),
            obj(76).ToString,
            obj(77).ToString,
            obj(78).ToString,
            obj(79).ToString,
            Convert.ToInt32(obj(80)),
            Convert.ToInt32(obj(81)),
            Convert.ToInt32(obj(82)),
            Convert.ToInt32(obj(83)),
            obj(84).ToString,
            obj(85).ToString,
            obj(86).ToString,
            obj(87).ToString,
            obj(88).ToString,
            obj(89).ToString,
            obj(90).ToString,
            obj(91).ToString,
            obj(92).ToString,
            obj(93).ToString,
            obj(94).ToString,
            obj(95).ToString,
            obj(96).ToString,
            obj(97).ToString,
            obj(98).ToString,
            Convert.ToInt32(obj(99)),
            obj(100).ToString,
            obj(101).ToString,
            obj(102).ToString,
            Convert.ToInt32(obj(103)),
            obj(104).ToString,
            obj(105).ToString,
            obj(106).ToString,
            obj(107).ToString,
            obj(108).ToString,
            obj(109).ToString)
        Next

        Return imctl

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