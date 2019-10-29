Imports System.ComponentModel
Imports System.Data.DataRowView
Imports System.Data.SqlClient
Imports System.Data
Imports System.Text
Imports System.Xml
Imports System.IO
Imports System.Data.SqlTypes


Public Class ItemLocationSetting
    Public loc As String
    Public alt1_type As String
    Public alt1_loc As String
    Public ri_mrb_loc As String
    Public loc_desc As String
    Public mult_bin_fg As String
    Public neg_inv_fg As String
    Public mrp_fg As String
    Public mn_no As String
    Public sb_no As String
    Public dp_no As String
    Public ri_rtv_mn_no As String
    Public ri_rtv_sb_no As String
    Public ri_rtv_dp_no As String
    Public ri_dit_mn_no As String
    Public ri_dit_sb_no As String
    Public ri_dit_dp_no As String
    Public scrap_mn_no As String
    Public scrap_sb_no As String
    Public scrap_dp_no As String
    Public accru_mn_no As String
    Public accru_sb_no As String
    Public accru_dp_no As String
    Public addr_1 As String
    Public addr_2 As String
    Public city As String
    Public state As String
    Public zip As String
    Public country As String
    Public phone_no As String
    Public phone_no_2 As String
    Public phone_ext_1 As String
    Public phone_ext_2 As String
    Public fax_no As String
    Public user_def_fld_1 As String
    Public user_def_fld_2 As String
    Public user_def_fld_3 As String
    Public user_def_fld_4 As String
    Public user_def_fld_5 As String
    Public frz_dt As Integer
    Public spc_fg As String
    Public web_item As String
    Public filler_0001 As String

    Const TABLE As String = "IMLOCFIL_SQL"

    Public Enum ItemLoc As Integer
        Factory = 1
        Warehouse = 2
    End Enum

    Private Function NewItemLocationSetting(ByVal m_loc As String, ByVal m_alt1_type As String, ByVal m_alt1_loc As String, ByVal m_ri_mrb_loc As String, ByVal m_loc_desc As String, _
                                            ByVal m_mult_bin_fg As String, ByVal m_neg_inv_fg As String, ByVal m_mrp_fg As String, ByVal m_mn_no As String, _
                                            ByVal m_sb_no As String, ByVal m_dp_no As String, ByVal m_ri_rtv_mn_no As String, ByVal m_ri_rtv_sb_no As String, ByVal m_ri_rtv_dp_no As String, _
                                            ByVal m_ri_dit_mn_no As String, ByVal m_ri_dit_sb_no As String, ByVal m_ri_dit_dp_no As String, ByVal m_scrap_mn_no As String, ByVal m_scrap_sb_no As String, _
                                            ByVal m_scrap_dp_no As String, ByVal m_accru_mn_no As String, ByVal m_accru_sb_no As String, ByVal m_accru_dp_no As String, ByVal m_addr_1 As String, _
                                            ByVal m_addr_2 As String, ByVal m_city As String, ByVal m_state As String, ByVal m_zip As String, ByVal m_country As String, _
                                            ByVal m_phone_no As String, ByVal m_phone_no_2 As String, ByVal m_phone_ext_1 As String, ByVal m_phone_ext_2 As String, ByVal m_fax_no As String, _
                                            ByVal m_user_def_fld_1 As String, ByVal m_user_def_fld_2 As String, ByVal m_user_def_fld_3 As String, ByVal m_user_def_fld_4 As String, ByVal m_user_def_fld_5 As String, ByVal m_frz_dt As Integer, _
                                            ByVal m_spc_fg As String, ByVal m_web_item As String, ByVal m_filler_0001 As String) As ItemLocationSetting




        Dim itmloc As New ItemLocationSetting
        itmloc.loc = m_loc
        itmloc.alt1_type = m_alt1_type
        itmloc.alt1_loc = m_alt1_loc
        itmloc.ri_mrb_loc = m_ri_mrb_loc
        itmloc.loc_desc = m_loc_desc
        itmloc.mult_bin_fg = m_mult_bin_fg
        itmloc.neg_inv_fg = m_neg_inv_fg
        itmloc.mrp_fg = m_mrp_fg
        itmloc.mn_no = m_mn_no
        itmloc.sb_no = m_sb_no
        itmloc.dp_no = m_dp_no
        itmloc.ri_rtv_mn_no = m_ri_rtv_mn_no
        itmloc.ri_rtv_sb_no = m_ri_rtv_sb_no
        itmloc.ri_rtv_dp_no = m_ri_rtv_dp_no
        itmloc.ri_dit_mn_no = m_ri_dit_mn_no
        itmloc.ri_dit_sb_no = m_ri_dit_sb_no
        itmloc.ri_dit_dp_no = m_ri_dit_dp_no
        itmloc.scrap_mn_no = m_scrap_mn_no
        itmloc.scrap_sb_no = m_scrap_sb_no
        itmloc.scrap_dp_no = m_scrap_dp_no
        itmloc.accru_mn_no = m_accru_mn_no
        itmloc.accru_sb_no = m_accru_sb_no
        itmloc.accru_dp_no = m_accru_dp_no
        itmloc.addr_1 = m_addr_1
        itmloc.addr_2 = m_addr_2
        itmloc.city = m_city
        itmloc.state = m_state
        itmloc.zip = m_zip
        itmloc.country = m_country
        itmloc.phone_no = m_phone_no
        itmloc.phone_no_2 = m_phone_no_2
        itmloc.phone_ext_1 = m_phone_ext_1
        itmloc.phone_ext_2 = m_phone_ext_2
        itmloc.fax_no = m_fax_no
        itmloc.user_def_fld_1 = m_user_def_fld_1
        itmloc.user_def_fld_2 = m_user_def_fld_2
        itmloc.user_def_fld_3 = m_user_def_fld_3
        itmloc.user_def_fld_4 = m_user_def_fld_4
        itmloc.user_def_fld_5 = m_user_def_fld_5
        itmloc.frz_dt = m_frz_dt
        itmloc.spc_fg = m_spc_fg
        itmloc.web_item = m_web_item
        itmloc.filler_0001 = m_filler_0001

        Return itmloc

    End Function



    Public Function GetItemLocationSetting(ByVal loc As String, ByVal cn As SqlConnection) As ItemLocationSetting

        'PRD_KEY_1 for Massarelli INVENTORY is "8"

        Dim itmloc As New ItemLocationSetting
        Dim dt As New DataTable
        Dim sSQL As String = "Select * from " & TABLE & " " _
                           & "Where loc = '" & loc & "' "

        dt = DAC.ExecuteSQL_DataTable(sSQL, cn, TABLE)

        itmloc = PopulateItemLocationSetting(dt)

        Return itmloc

    End Function

    Private Function PopulateItemLocationSetting(ByVal dt As DataTable) As ItemLocationSetting
        Dim itmloc As New ItemLocationSetting
        Dim obj(42) As Object
        For Each rw As DataRow In dt.Rows

            obj(0) = ReplaceNullValue(rw("loc"), dt.Columns("loc").DataType.ToString)
            obj(1) = ReplaceNullValue(rw("alt1_type"), dt.Columns("alt1_type").DataType.ToString)
            obj(2) = ReplaceNullValue(rw("alt1_loc"), dt.Columns("alt1_loc").DataType.ToString)
            obj(3) = ReplaceNullValue(rw("ri_mrb_loc"), dt.Columns("ri_mrb_loc").DataType.ToString)
            obj(4) = ReplaceNullValue(rw("loc_desc"), dt.Columns("loc_desc").DataType.ToString)
            obj(5) = ReplaceNullValue(rw("mult_bin_fg"), dt.Columns("mult_bin_fg").DataType.ToString)
            obj(6) = ReplaceNullValue(rw("neg_inv_fg"), dt.Columns("neg_inv_fg").DataType.ToString)
            obj(7) = ReplaceNullValue(rw("mrp_fg"), dt.Columns("mrp_fg").DataType.ToString)
            obj(8) = ReplaceNullValue(rw("mn_no"), dt.Columns("mn_no").DataType.ToString)
            obj(9) = ReplaceNullValue(rw("sb_no"), dt.Columns("sb_no").DataType.ToString)
            obj(10) = ReplaceNullValue(rw("dp_no"), dt.Columns("dp_no").DataType.ToString)
            obj(11) = ReplaceNullValue(rw("ri_rtv_mn_no"), dt.Columns("ri_rtv_mn_no").DataType.ToString)
            obj(12) = ReplaceNullValue(rw("ri_rtv_sb_no"), dt.Columns("ri_rtv_sb_no").DataType.ToString)
            obj(13) = ReplaceNullValue(rw("ri_rtv_dp_no"), dt.Columns("ri_rtv_dp_no").DataType.ToString)
            obj(14) = ReplaceNullValue(rw("ri_dit_mn_no"), dt.Columns("ri_dit_mn_no").DataType.ToString)
            obj(15) = ReplaceNullValue(rw("ri_dit_sb_no"), dt.Columns("ri_dit_sb_no").DataType.ToString)
            obj(16) = ReplaceNullValue(rw("ri_dit_dp_no"), dt.Columns("ri_dit_dp_no").DataType.ToString)
            obj(17) = ReplaceNullValue(rw("scrap_mn_no"), dt.Columns("scrap_mn_no").DataType.ToString)
            obj(18) = ReplaceNullValue(rw("scrap_sb_no"), dt.Columns("scrap_sb_no").DataType.ToString)
            obj(19) = ReplaceNullValue(rw("scrap_dp_no"), dt.Columns("scrap_dp_no").DataType.ToString)
            obj(20) = ReplaceNullValue(rw("accru_mn_no"), dt.Columns("accru_mn_no").DataType.ToString)
            obj(21) = ReplaceNullValue(rw("accru_sb_no"), dt.Columns("accru_sb_no").DataType.ToString)
            obj(22) = ReplaceNullValue(rw("accru_dp_no"), dt.Columns("accru_dp_no").DataType.ToString)
            obj(23) = ReplaceNullValue(rw("addr_1"), dt.Columns("addr_1").DataType.ToString)
            obj(24) = ReplaceNullValue(rw("addr_2"), dt.Columns("addr_2").DataType.ToString)
            obj(25) = ReplaceNullValue(rw("city"), dt.Columns("city").DataType.ToString)
            obj(26) = ReplaceNullValue(rw("state"), dt.Columns("state").DataType.ToString)
            obj(27) = ReplaceNullValue(rw("zip"), dt.Columns("zip").DataType.ToString)
            obj(28) = ReplaceNullValue(rw("country"), dt.Columns("country").DataType.ToString)
            obj(29) = ReplaceNullValue(rw("phone_no"), dt.Columns("phone_no").DataType.ToString)
            obj(30) = ReplaceNullValue(rw("phone_no_2"), dt.Columns("phone_no_2").DataType.ToString)
            obj(31) = ReplaceNullValue(rw("phone_ext_1"), dt.Columns("phone_ext_1").DataType.ToString)
            obj(32) = ReplaceNullValue(rw("phone_ext_2"), dt.Columns("phone_ext_2").DataType.ToString)
            obj(33) = ReplaceNullValue(rw("fax_no"), dt.Columns("fax_no").DataType.ToString)
            obj(34) = ReplaceNullValue(rw("user_def_fld_1"), dt.Columns("user_def_fld_1").DataType.ToString)
            obj(35) = ReplaceNullValue(rw("user_def_fld_2"), dt.Columns("user_def_fld_2").DataType.ToString)
            obj(36) = ReplaceNullValue(rw("user_def_fld_3"), dt.Columns("user_def_fld_3").DataType.ToString)
            obj(37) = ReplaceNullValue(rw("user_def_fld_4"), dt.Columns("user_def_fld_4").DataType.ToString)
            obj(38) = ReplaceNullValue(rw("user_def_fld_5"), dt.Columns("user_def_fld_5").DataType.ToString)
            obj(39) = ReplaceNullValue(rw("frz_dt"), dt.Columns("frz_dt").DataType.ToString)
            obj(40) = ReplaceNullValue(rw("spc_fg"), dt.Columns("spc_fg").DataType.ToString)
            obj(41) = ReplaceNullValue(rw("web_item"), dt.Columns("web_item").DataType.ToString)
            obj(42) = ReplaceNullValue(rw("filler_0001"), dt.Columns("filler_0001").DataType.ToString)

        Next

        itmloc = NewItemLocationSetting(
        obj(0).ToString,
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
        obj(11).ToString,
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
        Convert.ToInt32(obj(39)),
        obj(40).ToString,
        obj(41).ToString,
        obj(42).ToString)

        Return itmloc

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

    Public Function PadWarehouseLocation(ByVal loc As Integer) As String
        Return "00" + loc.ToString
    End Function
End Class
