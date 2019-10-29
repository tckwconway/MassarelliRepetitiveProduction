Imports System.Data.SqlClient

Public Class CompanySetting

    Public comp_key_1 As String
    Public rpt_name As String
    Public display_name As String
    Public addr_line_1 As String
    Public addr_line_2 As String
    Public addr_line_3 As String
    Public phone_no As String
    Public vers_of_sftwr As String
    Public uses_tts_fg As String
    Public gl_acct_lev_1_dgts As Integer
    Public gl_acct_lev_2_dgts As Integer
    Public gl_acct_lev_3_dgts As Integer
    Public uses_ar_fg As String
    Public uses_ar_drive As String
    Public uses_ap_fg As String
    Public uses_ap_drive As String
    Public uses_pr_fg As String
    Public uses_pr_drive As String
    Public uses_ad_fg As String
    Public uses_ad_drive As String
    Public uses_gl_fg As String
    Public uses_gl_drive As String
    Public uses_cm_fg As String
    Public uses_cm_drive As String
    Public uses_bb_fg As String
    Public uses_bb_drive As String
    Public uses_im_fg As String
    Public uses_im_drive As String
    Public uses_oe_fg As String
    Public uses_oe_drive As String
    Public uses_bomp_fg As String
    Public uses_bomp_drive As String
    Public uses_sfc_fg As String
    Public uses_sfc_drive As String
    Public uses_jc_fg As String
    Public uses_jc_drive As String
    Public uses_po_fg As String
    Public uses_po_drive As String
    Public uses_lp_fg As String
    Public uses_lp_drive As String
    Public uses_spc_fg As String
    Public uses_spc_drive As String
    Public uses_spr_fg As String
    Public uses_spr_drive As String
    Public uses_mrp_fg As String
    Public uses_mrp_drive As String
    Public uses_ms_fg As String
    Public uses_ms_drive As String
    Public uses_crp_fg As String
    Public uses_crp_drive As String
    Public has_mult_print As String
    Public spol_to_dsk_fg As String
    Public no_of_loc_prt As Integer
    Public no_of_net_prt As Integer
    Public spool_to_disk_def As String
    Public spool_to_ntwrk_def As String
    Public def_local_prter As Integer
    Public def_ntwrk_prter As Integer
    Public def_lpt1_prt_cd As Integer
    Public def_lpt2_prt_cd As Integer
    Public def_lpt3_prt_cd As Integer
    Public def_net0_prt_cd As Integer
    Public def_net1_prt_cd As Integer
    Public def_net2_prt_cd As Integer
    Public def_net3_prt_cd As Integer
    Public def_net4_prt_cd As Integer
    Public dt_format As String
    Public version_fg As String
    Public database_tp As String
    Public uses_mm_fg As String
    Public uses_mm_drive As String
    Public uses_sy_drive As String
    Public acct_audt_trail As String
    Public start_jnl_hist_no As Integer
    Public btrieve_owner_name As String
    Public comp_no_of_dec As Integer
    Public uses_pp_fg As String
    Public uses_pp_drive As String
    Public uses_rm_fg As String
    Public uses_rm_drive As String
    Public uses_de_fg As String
    Public uses_de_drive As String
    Public uses_je_fg As String
    Public uses_je_drive As String
    Public company_version As String
    Public purge_password As String
    Public phone_format As String
    Public rfc_no As String
    Public uses_ed_fg As String
    Public uses_ed_drive As String
    Public uses_bc_fg As String
    Public uses_bc_drive As String
    Public uses_mv_fg As String
    Public uses_mv_drive As String
    Public uses_sy_fg As String
    Public no_of_dec_cst_prc As Integer
    Public uses_tts_trx_fg As String
    Public hrs_decimal As Integer
    Public cst_decimal As Integer
    Public user_decimal As Integer
    Public rounding_fg As String
    Public rounding_mn_no As String
    Public rounding_sb_no As String
    Public rounding_dp_no As String
    Public local_kit_fg As String
    Public filler_0001 As String

    Const TABLE As String = "COMPFILE_SQL"


    Private Function NewCompanySetting(ByVal m_comp_key_1 As String, ByVal m_rpt_name As String, ByVal m_display_name As String, ByVal m_addr_line_1 As String, _
                                        ByVal m_addr_line_2 As String, ByVal m_addr_line_3 As String, ByVal m_phone_no As String, ByVal m_vers_of_sftwr As String, _
                                        ByVal m_uses_tts_fg As String, ByVal m_gl_acct_lev_1_dgts As Integer, ByVal m_gl_acct_lev_2_dgts As Integer, ByVal m_gl_acct_lev_3_dgts As Integer, _
                                        ByVal m_uses_ar_fg As String, ByVal m_uses_ar_drive As String, ByVal m_uses_ap_fg As String, ByVal m_uses_ap_drive As String, ByVal m_uses_pr_fg As String, _
                                        ByVal m_uses_pr_drive As String, ByVal m_uses_ad_fg As String, ByVal m_uses_ad_drive As String, ByVal m_uses_gl_fg As String, ByVal m_uses_gl_drive As String, _
                                        ByVal m_uses_cm_fg As String, ByVal m_uses_cm_drive As String, ByVal m_uses_bb_fg As String, ByVal m_uses_bb_drive As String, ByVal m_uses_im_fg As String, _
                                        ByVal m_uses_im_drive As String, ByVal m_uses_oe_fg As String, ByVal m_uses_oe_drive As String, ByVal m_uses_bomp_fg As String, ByVal m_uses_bomp_drive As String, _
                                        ByVal m_uses_sfc_fg As String, ByVal m_uses_sfc_drive As String, ByVal m_uses_jc_fg As String, ByVal m_uses_jc_drive As String, ByVal m_uses_po_fg As String, _
                                        ByVal m_uses_po_drive As String, ByVal m_uses_lp_fg As String, ByVal m_uses_lp_drive As String, ByVal m_uses_spc_fg As String, ByVal m_uses_spc_drive As String, _
                                        ByVal m_uses_spr_fg As String, ByVal m_uses_spr_drive As String, ByVal m_uses_mrp_fg As String, ByVal m_uses_mrp_drive As String, ByVal m_uses_ms_fg As String, _
                                        ByVal m_uses_ms_drive As String, ByVal m_uses_crp_fg As String, ByVal m_uses_crp_drive As String, ByVal m_has_mult_print As String, ByVal m_spol_to_dsk_fg As String, _
                                        ByVal m_no_of_loc_prt As Integer, ByVal m_no_of_net_prt As Integer, ByVal m_spool_to_disk_def As String, ByVal m_spool_to_ntwrk_def As String, ByVal m_def_local_prter As Integer, _
                                        ByVal m_def_ntwrk_prter As Integer, ByVal m_def_lpt1_prt_cd As Integer, ByVal m_def_lpt2_prt_cd As Integer, ByVal m_def_lpt3_prt_cd As Integer, _
                                        ByVal m_def_net0_prt_cd As Integer, ByVal m_def_net1_prt_cd As Integer, ByVal m_def_net2_prt_cd As Integer, ByVal m_def_net3_prt_cd As Integer, _
                                        ByVal m_def_net4_prt_cd As Integer, ByVal m_dt_format As String, ByVal m_version_fg As String, ByVal m_database_tp As String, ByVal m_uses_mm_fg As String, _
                                        ByVal m_uses_mm_drive As String, ByVal m_uses_sy_drive As String, ByVal m_acct_audt_trail As String, ByVal m_start_jnl_hist_no As Integer, ByVal m_btrieve_owner_name As String, _
                                        ByVal m_comp_no_of_dec As Integer, ByVal m_uses_pp_fg As String, ByVal m_uses_pp_drive As String, ByVal m_uses_rm_fg As String, ByVal m_uses_rm_drive As String, _
                                        ByVal m_uses_de_fg As String, ByVal m_uses_de_drive As String, ByVal m_uses_je_fg As String, ByVal m_uses_je_drive As String, ByVal m_company_version As String, _
                                        ByVal m_purge_password As String, ByVal m_phone_format As String, ByVal m_rfc_no As String, ByVal m_uses_ed_fg As String, ByVal m_uses_ed_drive As String, _
                                        ByVal m_uses_bc_fg As String, ByVal m_uses_bc_drive As String, ByVal m_uses_mv_fg As String, ByVal m_uses_mv_drive As String, ByVal m_uses_sy_fg As String, _
                                        ByVal m_no_of_dec_cst_prc As Integer, ByVal m_uses_tts_trx_fg As String, ByVal m_hrs_decimal As Integer, ByVal m_cst_decimal As Integer, ByVal m_user_decimal As Integer, _
                                        ByVal m_rounding_fg As String, ByVal m_rounding_mn_no As String, ByVal m_rounding_sb_no As String, ByVal m_rounding_dp_no As String, ByVal m_local_kit_fg As String, _
                                        ByVal m_filler_0001 As String) As CompanySetting

        Dim comp As New CompanySetting
        comp.comp_key_1 = m_comp_key_1
        comp.rpt_name = m_rpt_name
        comp.display_name = m_display_name
        comp.addr_line_1 = m_addr_line_1
        comp.addr_line_2 = m_addr_line_2
        comp.addr_line_3 = m_addr_line_3
        comp.phone_no = m_phone_no
        comp.vers_of_sftwr = m_vers_of_sftwr
        comp.uses_tts_fg = m_uses_tts_fg
        comp.gl_acct_lev_1_dgts = m_gl_acct_lev_1_dgts
        comp.gl_acct_lev_2_dgts = m_gl_acct_lev_2_dgts
        comp.gl_acct_lev_3_dgts = m_gl_acct_lev_3_dgts
        comp.uses_ar_fg = m_uses_ar_fg
        comp.uses_ar_drive = m_uses_ar_drive
        comp.uses_ap_fg = m_uses_ap_fg
        comp.uses_ap_drive = m_uses_ap_drive
        comp.uses_pr_fg = m_uses_pr_fg
        comp.uses_pr_drive = m_uses_pr_drive
        comp.uses_ad_fg = m_uses_ad_fg
        comp.uses_ad_drive = m_uses_ad_drive
        comp.uses_gl_fg = m_uses_gl_fg
        comp.uses_gl_drive = m_uses_gl_drive
        comp.uses_cm_fg = m_uses_cm_fg
        comp.uses_cm_drive = m_uses_cm_drive
        comp.uses_bb_fg = m_uses_bb_fg
        comp.uses_bb_drive = m_uses_bb_drive
        comp.uses_im_fg = m_uses_im_fg
        comp.uses_im_drive = m_uses_im_drive
        comp.uses_oe_fg = m_uses_oe_fg
        comp.uses_oe_drive = m_uses_oe_drive
        comp.uses_bomp_fg = m_uses_bomp_fg
        comp.uses_bomp_drive = m_uses_bomp_drive
        comp.uses_sfc_fg = m_uses_sfc_fg
        comp.uses_sfc_drive = m_uses_sfc_drive
        comp.uses_jc_fg = m_uses_jc_fg
        comp.uses_jc_drive = m_uses_jc_drive
        comp.uses_po_fg = m_uses_po_fg
        comp.uses_po_drive = m_uses_po_drive
        comp.uses_lp_fg = m_uses_lp_fg
        comp.uses_lp_drive = m_uses_lp_drive
        comp.uses_spc_fg = m_uses_spc_fg
        comp.uses_spc_drive = m_uses_spc_drive
        comp.uses_spr_fg = m_uses_spr_fg
        comp.uses_spr_drive = m_uses_spr_drive
        comp.uses_mrp_fg = m_uses_mrp_fg
        comp.uses_mrp_drive = m_uses_mrp_drive
        comp.uses_ms_fg = m_uses_ms_fg
        comp.uses_ms_drive = m_uses_ms_drive
        comp.uses_crp_fg = m_uses_crp_fg
        comp.uses_crp_drive = m_uses_crp_drive
        comp.has_mult_print = m_has_mult_print
        comp.spol_to_dsk_fg = m_spol_to_dsk_fg
        comp.no_of_loc_prt = m_no_of_loc_prt
        comp.no_of_net_prt = m_no_of_net_prt
        comp.spool_to_disk_def = m_spool_to_disk_def
        comp.spool_to_ntwrk_def = m_spool_to_ntwrk_def
        comp.def_local_prter = m_def_local_prter
        comp.def_ntwrk_prter = m_def_ntwrk_prter
        comp.def_lpt1_prt_cd = m_def_lpt1_prt_cd
        comp.def_lpt2_prt_cd = m_def_lpt2_prt_cd
        comp.def_lpt3_prt_cd = m_def_lpt3_prt_cd
        comp.def_net0_prt_cd = m_def_net0_prt_cd
        comp.def_net1_prt_cd = m_def_net1_prt_cd
        comp.def_net2_prt_cd = m_def_net2_prt_cd
        comp.def_net3_prt_cd = m_def_net3_prt_cd
        comp.def_net4_prt_cd = m_def_net4_prt_cd
        comp.dt_format = m_dt_format
        comp.version_fg = m_version_fg
        comp.database_tp = m_database_tp
        comp.uses_mm_fg = m_uses_mm_fg
        comp.uses_mm_drive = m_uses_mm_drive
        comp.uses_sy_drive = m_uses_sy_drive
        comp.acct_audt_trail = m_acct_audt_trail
        comp.start_jnl_hist_no = m_start_jnl_hist_no
        comp.btrieve_owner_name = m_btrieve_owner_name
        comp.comp_no_of_dec = m_comp_no_of_dec
        comp.uses_pp_fg = m_uses_pp_fg
        comp.uses_pp_drive = m_uses_pp_drive
        comp.uses_rm_fg = m_uses_rm_fg
        comp.uses_rm_drive = m_uses_rm_drive
        comp.uses_de_fg = m_uses_de_fg
        comp.uses_de_drive = m_uses_de_drive
        comp.uses_je_fg = m_uses_je_fg
        comp.uses_je_drive = m_uses_je_drive
        comp.company_version = m_company_version
        comp.purge_password = m_purge_password
        comp.phone_format = m_phone_format
        comp.rfc_no = m_rfc_no
        comp.uses_ed_fg = m_uses_ed_fg
        comp.uses_ed_drive = m_uses_ed_drive
        comp.uses_bc_fg = m_uses_bc_fg
        comp.uses_bc_drive = m_uses_bc_drive
        comp.uses_mv_fg = m_uses_mv_fg
        comp.uses_mv_drive = m_uses_mv_drive
        comp.uses_sy_fg = m_uses_sy_fg
        comp.no_of_dec_cst_prc = m_no_of_dec_cst_prc
        comp.uses_tts_trx_fg = m_uses_tts_trx_fg
        comp.hrs_decimal = m_hrs_decimal
        comp.cst_decimal = m_cst_decimal
        comp.user_decimal = m_user_decimal
        comp.rounding_fg = m_rounding_fg
        comp.rounding_mn_no = m_rounding_mn_no
        comp.rounding_sb_no = m_rounding_sb_no
        comp.rounding_dp_no = m_rounding_dp_no
        comp.local_kit_fg = m_local_kit_fg
        comp.filler_0001 = m_filler_0001

        Return comp

    End Function


    Public Function GetCompanySetting(ByVal comp_key As String, ByVal cn As SqlConnection) As CompanySetting

        'Company comp_key_1 for Massarelli is "1"

        Dim comp As New CompanySetting
        Dim dt As New DataTable
        Dim sSQL As String = "Select * from " & TABLE & " " _
                           & "Where comp_key_1 = '" & comp_key & "' "

        dt = DAC.ExecuteSQL_DataTable(sSQL, cn, TABLE)

        comp = PopulateCompanySettings(dt)

        Return comp

    End Function

    Private Function PopulateCompanySettings(ByVal dt As DataTable) As CompanySetting
        Dim comp As New CompanySetting
        Dim obj(105) As Object
        For Each rw As DataRow In dt.Rows

            obj(0) = ReplaceNullValue(rw("comp_key_1"), dt.Columns("comp_key_1").DataType.ToString)
            obj(1) = ReplaceNullValue(rw("rpt_name"), dt.Columns("rpt_name").DataType.ToString)
            obj(2) = ReplaceNullValue(rw("display_name"), dt.Columns("display_name").DataType.ToString)
            obj(3) = ReplaceNullValue(rw("addr_line_1"), dt.Columns("addr_line_1").DataType.ToString)
            obj(4) = ReplaceNullValue(rw("addr_line_2"), dt.Columns("addr_line_2").DataType.ToString)
            obj(5) = ReplaceNullValue(rw("addr_line_3"), dt.Columns("addr_line_3").DataType.ToString)
            obj(6) = ReplaceNullValue(rw("phone_no"), dt.Columns("phone_no").DataType.ToString)
            obj(7) = ReplaceNullValue(rw("vers_of_sftwr"), dt.Columns("vers_of_sftwr").DataType.ToString)
            obj(8) = ReplaceNullValue(rw("uses_tts_fg"), dt.Columns("uses_tts_fg").DataType.ToString)
            obj(9) = ReplaceNullValue(rw("gl_acct_lev_1_dgts"), dt.Columns("gl_acct_lev_1_dgts").DataType.ToString)
            obj(10) = ReplaceNullValue(rw("gl_acct_lev_2_dgts"), dt.Columns("gl_acct_lev_2_dgts").DataType.ToString)
            obj(11) = ReplaceNullValue(rw("gl_acct_lev_3_dgts"), dt.Columns("gl_acct_lev_3_dgts").DataType.ToString)
            obj(12) = ReplaceNullValue(rw("uses_ar_fg"), dt.Columns("uses_ar_fg").DataType.ToString)
            obj(13) = ReplaceNullValue(rw("uses_ar_drive"), dt.Columns("uses_ar_drive").DataType.ToString)
            obj(14) = ReplaceNullValue(rw("uses_ap_fg"), dt.Columns("uses_ap_fg").DataType.ToString)
            obj(15) = ReplaceNullValue(rw("uses_ap_drive"), dt.Columns("uses_ap_drive").DataType.ToString)
            obj(16) = ReplaceNullValue(rw("uses_pr_fg"), dt.Columns("uses_pr_fg").DataType.ToString)
            obj(17) = ReplaceNullValue(rw("uses_pr_drive"), dt.Columns("uses_pr_drive").DataType.ToString)
            obj(18) = ReplaceNullValue(rw("uses_ad_fg"), dt.Columns("uses_ad_fg").DataType.ToString)
            obj(19) = ReplaceNullValue(rw("uses_ad_drive"), dt.Columns("uses_ad_drive").DataType.ToString)
            obj(20) = ReplaceNullValue(rw("uses_gl_fg"), dt.Columns("uses_gl_fg").DataType.ToString)
            obj(21) = ReplaceNullValue(rw("uses_gl_drive"), dt.Columns("uses_gl_drive").DataType.ToString)
            obj(22) = ReplaceNullValue(rw("uses_cm_fg"), dt.Columns("uses_cm_fg").DataType.ToString)
            obj(23) = ReplaceNullValue(rw("uses_cm_drive"), dt.Columns("uses_cm_drive").DataType.ToString)
            obj(24) = ReplaceNullValue(rw("uses_bb_fg"), dt.Columns("uses_bb_fg").DataType.ToString)
            obj(25) = ReplaceNullValue(rw("uses_bb_drive"), dt.Columns("uses_bb_drive").DataType.ToString)
            obj(26) = ReplaceNullValue(rw("uses_im_fg"), dt.Columns("uses_im_fg").DataType.ToString)
            obj(27) = ReplaceNullValue(rw("uses_im_drive"), dt.Columns("uses_im_drive").DataType.ToString)
            obj(28) = ReplaceNullValue(rw("uses_oe_fg"), dt.Columns("uses_oe_fg").DataType.ToString)
            obj(29) = ReplaceNullValue(rw("uses_oe_drive"), dt.Columns("uses_oe_drive").DataType.ToString)
            obj(30) = ReplaceNullValue(rw("uses_bomp_fg"), dt.Columns("uses_bomp_fg").DataType.ToString)
            obj(31) = ReplaceNullValue(rw("uses_bomp_drive"), dt.Columns("uses_bomp_drive").DataType.ToString)
            obj(32) = ReplaceNullValue(rw("uses_sfc_fg"), dt.Columns("uses_sfc_fg").DataType.ToString)
            obj(33) = ReplaceNullValue(rw("uses_sfc_drive"), dt.Columns("uses_sfc_drive").DataType.ToString)
            obj(34) = ReplaceNullValue(rw("uses_jc_fg"), dt.Columns("uses_jc_fg").DataType.ToString)
            obj(35) = ReplaceNullValue(rw("uses_jc_drive"), dt.Columns("uses_jc_drive").DataType.ToString)
            obj(36) = ReplaceNullValue(rw("uses_po_fg"), dt.Columns("uses_po_fg").DataType.ToString)
            obj(37) = ReplaceNullValue(rw("uses_po_drive"), dt.Columns("uses_po_drive").DataType.ToString)
            obj(38) = ReplaceNullValue(rw("uses_lp_fg"), dt.Columns("uses_lp_fg").DataType.ToString)
            obj(39) = ReplaceNullValue(rw("uses_lp_drive"), dt.Columns("uses_lp_drive").DataType.ToString)
            obj(40) = ReplaceNullValue(rw("uses_spc_fg"), dt.Columns("uses_spc_fg").DataType.ToString)
            obj(41) = ReplaceNullValue(rw("uses_spc_drive"), dt.Columns("uses_spc_drive").DataType.ToString)
            obj(42) = ReplaceNullValue(rw("uses_spr_fg"), dt.Columns("uses_spr_fg").DataType.ToString)
            obj(43) = ReplaceNullValue(rw("uses_spr_drive"), dt.Columns("uses_spr_drive").DataType.ToString)
            obj(44) = ReplaceNullValue(rw("uses_mrp_fg"), dt.Columns("uses_mrp_fg").DataType.ToString)
            obj(45) = ReplaceNullValue(rw("uses_mrp_drive"), dt.Columns("uses_mrp_drive").DataType.ToString)
            obj(46) = ReplaceNullValue(rw("uses_ms_fg"), dt.Columns("uses_ms_fg").DataType.ToString)
            obj(47) = ReplaceNullValue(rw("uses_ms_drive"), dt.Columns("uses_ms_drive").DataType.ToString)
            obj(48) = ReplaceNullValue(rw("uses_crp_fg"), dt.Columns("uses_crp_fg").DataType.ToString)
            obj(49) = ReplaceNullValue(rw("uses_crp_drive"), dt.Columns("uses_crp_drive").DataType.ToString)
            obj(50) = ReplaceNullValue(rw("has_mult_print"), dt.Columns("has_mult_print").DataType.ToString)
            obj(51) = ReplaceNullValue(rw("spol_to_dsk_fg"), dt.Columns("spol_to_dsk_fg").DataType.ToString)
            obj(52) = ReplaceNullValue(rw("no_of_loc_prt"), dt.Columns("no_of_loc_prt").DataType.ToString)
            obj(53) = ReplaceNullValue(rw("no_of_net_prt"), dt.Columns("no_of_net_prt").DataType.ToString)
            obj(54) = ReplaceNullValue(rw("spool_to_disk_def"), dt.Columns("spool_to_disk_def").DataType.ToString)
            obj(55) = ReplaceNullValue(rw("spool_to_ntwrk_def"), dt.Columns("spool_to_ntwrk_def").DataType.ToString)
            obj(56) = ReplaceNullValue(rw("def_local_prter"), dt.Columns("def_local_prter").DataType.ToString)
            obj(57) = ReplaceNullValue(rw("def_ntwrk_prter"), dt.Columns("def_ntwrk_prter").DataType.ToString)
            obj(58) = ReplaceNullValue(rw("def_lpt1_prt_cd"), dt.Columns("def_lpt1_prt_cd").DataType.ToString)
            obj(59) = ReplaceNullValue(rw("def_lpt2_prt_cd"), dt.Columns("def_lpt2_prt_cd").DataType.ToString)
            obj(60) = ReplaceNullValue(rw("def_lpt3_prt_cd"), dt.Columns("def_lpt3_prt_cd").DataType.ToString)
            obj(61) = ReplaceNullValue(rw("def_net0_prt_cd"), dt.Columns("def_net0_prt_cd").DataType.ToString)
            obj(62) = ReplaceNullValue(rw("def_net1_prt_cd"), dt.Columns("def_net1_prt_cd").DataType.ToString)
            obj(63) = ReplaceNullValue(rw("def_net2_prt_cd"), dt.Columns("def_net2_prt_cd").DataType.ToString)
            obj(64) = ReplaceNullValue(rw("def_net3_prt_cd"), dt.Columns("def_net3_prt_cd").DataType.ToString)
            obj(65) = ReplaceNullValue(rw("def_net4_prt_cd"), dt.Columns("def_net4_prt_cd").DataType.ToString)
            obj(66) = ReplaceNullValue(rw("dt_format"), dt.Columns("dt_format").DataType.ToString)
            obj(67) = ReplaceNullValue(rw("version_fg"), dt.Columns("version_fg").DataType.ToString)
            obj(68) = ReplaceNullValue(rw("database_tp"), dt.Columns("database_tp").DataType.ToString)
            obj(69) = ReplaceNullValue(rw("uses_mm_fg"), dt.Columns("uses_mm_fg").DataType.ToString)
            obj(70) = ReplaceNullValue(rw("uses_mm_drive"), dt.Columns("uses_mm_drive").DataType.ToString)
            obj(71) = ReplaceNullValue(rw("uses_sy_drive"), dt.Columns("uses_sy_drive").DataType.ToString)
            obj(72) = ReplaceNullValue(rw("acct_audt_trail"), dt.Columns("acct_audt_trail").DataType.ToString)
            obj(73) = ReplaceNullValue(rw("start_jnl_hist_no"), dt.Columns("start_jnl_hist_no").DataType.ToString)
            obj(74) = ReplaceNullValue(rw("btrieve_owner_name"), dt.Columns("btrieve_owner_name").DataType.ToString)
            obj(75) = ReplaceNullValue(rw("comp_no_of_dec"), dt.Columns("comp_no_of_dec").DataType.ToString)
            obj(76) = ReplaceNullValue(rw("uses_pp_fg"), dt.Columns("uses_pp_fg").DataType.ToString)
            obj(77) = ReplaceNullValue(rw("uses_pp_drive"), dt.Columns("uses_pp_drive").DataType.ToString)
            obj(78) = ReplaceNullValue(rw("uses_rm_fg"), dt.Columns("uses_rm_fg").DataType.ToString)
            obj(79) = ReplaceNullValue(rw("uses_rm_drive"), dt.Columns("uses_rm_drive").DataType.ToString)
            obj(80) = ReplaceNullValue(rw("uses_de_fg"), dt.Columns("uses_de_fg").DataType.ToString)
            obj(81) = ReplaceNullValue(rw("uses_de_drive"), dt.Columns("uses_de_drive").DataType.ToString)
            obj(82) = ReplaceNullValue(rw("uses_je_fg"), dt.Columns("uses_je_fg").DataType.ToString)
            obj(83) = ReplaceNullValue(rw("uses_je_drive"), dt.Columns("uses_je_drive").DataType.ToString)
            obj(84) = ReplaceNullValue(rw("company_version"), dt.Columns("company_version").DataType.ToString)
            obj(85) = ReplaceNullValue(rw("purge_password"), dt.Columns("purge_password").DataType.ToString)
            obj(86) = ReplaceNullValue(rw("phone_format"), dt.Columns("phone_format").DataType.ToString)
            obj(87) = ReplaceNullValue(rw("rfc_no"), dt.Columns("rfc_no").DataType.ToString)
            obj(88) = ReplaceNullValue(rw("uses_ed_fg"), dt.Columns("uses_ed_fg").DataType.ToString)
            obj(89) = ReplaceNullValue(rw("uses_ed_drive"), dt.Columns("uses_ed_drive").DataType.ToString)
            obj(90) = ReplaceNullValue(rw("uses_bc_fg"), dt.Columns("uses_bc_fg").DataType.ToString)
            obj(91) = ReplaceNullValue(rw("uses_bc_drive"), dt.Columns("uses_bc_drive").DataType.ToString)
            obj(92) = ReplaceNullValue(rw("uses_mv_fg"), dt.Columns("uses_mv_fg").DataType.ToString)
            obj(93) = ReplaceNullValue(rw("uses_mv_drive"), dt.Columns("uses_mv_drive").DataType.ToString)
            obj(94) = ReplaceNullValue(rw("uses_sy_fg"), dt.Columns("uses_sy_fg").DataType.ToString)
            obj(95) = ReplaceNullValue(rw("no_of_dec_cst_prc"), dt.Columns("no_of_dec_cst_prc").DataType.ToString)
            obj(96) = ReplaceNullValue(rw("uses_tts_trx_fg"), dt.Columns("uses_tts_trx_fg").DataType.ToString)
            obj(97) = ReplaceNullValue(rw("hrs_decimal"), dt.Columns("hrs_decimal").DataType.ToString)
            obj(98) = ReplaceNullValue(rw("cst_decimal"), dt.Columns("cst_decimal").DataType.ToString)
            obj(99) = ReplaceNullValue(rw("user_decimal"), dt.Columns("user_decimal").DataType.ToString)
            obj(100) = ReplaceNullValue(rw("rounding_fg"), dt.Columns("rounding_fg").DataType.ToString)
            obj(101) = ReplaceNullValue(rw("rounding_mn_no"), dt.Columns("rounding_mn_no").DataType.ToString)
            obj(102) = ReplaceNullValue(rw("rounding_sb_no"), dt.Columns("rounding_sb_no").DataType.ToString)
            obj(103) = ReplaceNullValue(rw("rounding_dp_no"), dt.Columns("rounding_dp_no").DataType.ToString)
            obj(104) = ReplaceNullValue(rw("local_kit_fg"), dt.Columns("local_kit_fg").DataType.ToString)
            obj(105) = ReplaceNullValue(rw("filler_0001"), dt.Columns("filler_0001").DataType.ToString)

            comp = NewCompanySetting(
            obj(0).ToString,
            obj(1).ToString,
            obj(2).ToString,
            obj(3).ToString,
            obj(4).ToString,
            obj(5).ToString,
            obj(6).ToString,
            obj(7).ToString,
            obj(8).ToString,
            Convert.ToInt16(obj(9)),
            Convert.ToInt16(obj(10)),
            Convert.ToInt16(obj(11)),
            obj(12).ToString,
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
            Convert.ToInt16(obj(52)),
            Convert.ToInt16(obj(53)),
            obj(54).ToString,
            obj(55).ToString,
            Convert.ToInt16(obj(56)),
            Convert.ToInt16(obj(57)),
            Convert.ToInt16(obj(58)),
            Convert.ToInt16(obj(59)),
            Convert.ToInt16(obj(60)),
            Convert.ToInt16(obj(61)),
            Convert.ToInt16(obj(62)),
            Convert.ToInt16(obj(63)),
            Convert.ToInt16(obj(64)),
            Convert.ToInt16(obj(65)),
            obj(66).ToString,
            obj(67).ToString,
            obj(68).ToString,
            obj(69).ToString,
            obj(70).ToString,
            obj(71).ToString,
            obj(72).ToString,
            Convert.ToInt16(obj(73)),
            obj(74).ToString,
            Convert.ToInt16(obj(75)),
            obj(76).ToString,
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
            obj(92).ToString,
            obj(93).ToString,
            obj(94).ToString,
            Convert.ToInt16(obj(95)),
            obj(96).ToString,
            Convert.ToInt16(obj(97)),
            Convert.ToInt16(obj(98)),
            Convert.ToInt16(obj(99)),
            obj(100).ToString,
            obj(101).ToString,
            obj(102).ToString,
            obj(103).ToString,
            obj(104).ToString,
            obj(105).ToString)
        Next

        Return comp

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

            Case Else
                Return Nothing

        End Select
    End Function
End Class
