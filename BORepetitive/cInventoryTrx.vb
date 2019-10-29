Imports System.ComponentModel
Imports System.Data.DataRowView
Imports System.Data.SqlClient
Imports System.Data
Imports System.Text
Imports System.Xml
Imports System.IO
Imports System.Data.SqlTypes

Public Class InventoryTrx
    Public source As String
    Public ord_no As String
    Public ctl_no As Integer
    Public line_no As Integer
    Public lev_no As Integer
    Public seq_no As Integer
    Public from_source As String
    Public from_ord_no As String
    Public from_ctl_no As Integer
    Public from_line_no As Integer
    Public from_lev_no As Integer
    Public from_seq_no As Integer
    Public item_no As String
    Public item_filler As String
    Public par_item_no As String
    Public par_item_filler As String
    Public loc As String
    Public trx_dt As Integer
    Public trx_tm As Integer
    Public doc_dt As Integer
    Public doc_type As String
    Public doc_ord_no As String
    Public doc_source As String
    Public cus_no As String
    Public vend_no As String
    Public prod_type As String
    Public quantity As Integer
    Public old_quantity As Integer
    Public unit_cost As Integer
    Public old_unit_cost As Integer
    Public new_unit_cost As Integer
    Public price As Integer
    Public build_qty As Integer
    Public build_qty_per As Integer
    Public amt As Integer
    Public landed_cost As Integer
    Public receipt_ord_no As String
    Public status As String
    Public jnl As String
    Public batch_id As String
    Public user_name As String
    Public id_no As String
    Public comment As String
    Public filler_0003 As String
    Public trx_qty_bkord As Integer
    Public promise_dt As Integer
    Public rev_no As String
    Public deall_amt As Integer
    Public filler_0004 As String


    Const TABLE As String = "IMINVTRX_SQL"

    Public Shared Function NewInventoryTrx(tranData As Object, tableData As Object, cn As SqlConnection) As InventoryTrx

        Dim invTrx As New InventoryTrx
        Dim doctype As String = ""
        If tranData.TranType = "Transfer" Then
            doctype = "T"
        ElseIf tranData.TranType = "Issue" Then
            doctype = "I"
        ElseIf tranData.TranType = "Receipt" Then
            doctype = "R"
        End If


        With tranData
            invTrx.source = "I"
            invTrx.ord_no = .OrderNo
            invTrx.ctl_no = 0
            invTrx.line_no = 0
            invTrx.lev_no = .LevelNo
            invTrx.seq_no = 0
            invTrx.from_source = Nothing
            invTrx.from_ord_no = Nothing
            invTrx.from_ctl_no = 0
            invTrx.from_line_no = 0
            invTrx.from_lev_no = 0
            invTrx.from_seq_no = 0
            invTrx.item_no = .ItemNo
            invTrx.item_filler = ""
            invTrx.par_item_no = ""
            invTrx.par_item_filler = ""
            invTrx.loc = .Loc
            invTrx.trx_dt = cBusObj.GetMacolaDate(Now)
            invTrx.trx_tm = cBusObj.GetMacolaTime(Now)
            invTrx.doc_dt = tranData.TransEffectiveDate
            invTrx.doc_type = doctype
            invTrx.doc_ord_no = Nothing
            invTrx.doc_source = Nothing
            invTrx.cus_no = Nothing
            invTrx.vend_no = Nothing
            invTrx.prod_type = Nothing
            invTrx.quantity = .qtytomove
            'HOW old_quantity IS CALCULATED?:  It's not dumbell, its just taken from IMINVLOC_SQL qty_on_hand before it's updated, that's all.  
            Dim ssql As String = "Select qty_on_hand from IMINVLOC_SQL where item_no = '" & tranData.ItemNo.ToString.Trim & "' and loc = '" & tranData.Loc.ToString.Trim & "'"
            invTrx.old_quantity = Convert.ToDecimal(cBusObj.ExecuteSQLScalar(ssql, cn))
            ' ''FORGET ALL THIS .... HOW old_quantity IS CALCULATED: Calculation is done using values from the last ININVTRX record for the item_no and Loc, 
            ' ''                                If the last doc_type was I (Issue), then, retrieve the last transaction & subtract the last Tran quantity from _old_quantity
            ' ''                                and if it's R, add those values together.  BUT IT USES THE LAST TRANSACTION, not the current values,  
            ' ''                                and carries the sum or difference of those values forward into the next transaction as the current old_quantity. 
            ' ''                                For T (transfers) if lev_no = 0 then old_quantity - quantity, if lev_no = 1 then old_quantity + quantity. 
            ''If tableData.doc_type = "R" Then
            ''    invTrx.old_quantity = tableData.old_quantity + tableData.quantity
            ''ElseIf tableData.doc_type = "I" Then
            ''    invTrx.old_quantity = tableData.old_quantity + tableData.quantity * -1
            ''ElseIf tableData.doc_type = "T" And tableData.lev_no = 0 Then
            ''    'update_iminv is set to N when transfers occur within the same warehouse, thus, the old_quantity does not change ...
            ''    If tranData.UpdateIMINV = "N" Then
            ''        invTrx.old_quantity = tableData.quantity
            ''    Else
            ''        invTrx.old_quantity = tableData.old_quantity + tableData.quantity * -1
            ''    End If
            ''ElseIf tableData.doc_type = "T" And tableData.lev_no = 1 Then
            ''    If tranData.UpdateIMINV = "N" Then
            ''        invTrx.old_quantity = tableData.quantity
            ''    Else
            ''        invTrx.old_quantity = tableData.old_quantity + tableData.quantity
            ''    End If
            ''End If
            invTrx.unit_cost = 0
            invTrx.old_unit_cost = 0
            invTrx.new_unit_cost = 0
            invTrx.price = 0
            invTrx.build_qty = 0
            invTrx.build_qty_per = 0
            invTrx.amt = 0
            invTrx.landed_cost = 0
            invTrx.receipt_ord_no = 0
            invTrx.status = Nothing
            invTrx.jnl = ""
            invTrx.batch_id = ""
            invTrx.user_name = .UserName
            invTrx.id_no = .IDNo
            invTrx.comment = Nothing
            invTrx.filler_0003 = Nothing
            invTrx.trx_qty_bkord = 0
            invTrx.promise_dt = tranData.TransEffectiveDate
            invTrx.rev_no = Nothing
            invTrx.deall_amt = 0
            invTrx.filler_0004 = Nothing

        End With

        Return invTrx

    End Function


    Public Shared Function NewInventoryTrx(ByVal msource As String, ByVal mord_no As String, ByVal mctl_no As Integer, ByVal mline_no As Integer, _
                                           ByVal mlev_no As Integer, ByVal mseq_no As Integer, ByVal mfrom_source As String, ByVal mfrom_ord_no As String, _
                                           ByVal mfrom_ctl_no As Integer, ByVal mfrom_line_no As Integer, ByVal mfrom_lev_no As Integer, _
                                           ByVal mfrom_seq_no As Integer, ByVal mitem_no As String, ByVal mitem_filler As String, _
                                           ByVal mpar_item_no As String, ByVal mpar_item_filler As String, ByVal mloc As String, _
                                           ByVal mtrx_dt As Integer, ByVal mtrx_tm As Integer, ByVal mdoc_dt As Integer, _
                                           ByVal mdoc_type As String, ByVal mdoc_ord_no As String, ByVal mdoc_source As String, _
                                           ByVal mcus_no As String, ByVal mvend_no As String, ByVal mprod_type As String, _
                                           ByVal mquantity As Decimal, ByVal mold_quantity As Decimal, ByVal munit_cost As Decimal, _
                                           ByVal mold_unit_cost As Decimal, ByVal mnew_unit_cost As Decimal, ByVal mprice As Decimal, _
                                           ByVal mbuild_qty As Decimal, ByVal mbuild_qty_per As Decimal, ByVal mamt As Decimal, _
                                           ByVal mlanded_cost As Decimal, ByVal mreceipt_ord_no As String, ByVal mstatus As String, _
                                           ByVal mjnl As String, ByVal mbatch_id As String, ByVal muser_name As String, ByVal mid_no As String, _
                                           ByVal mcomment As String, ByVal mfiller_0003 As String, ByVal mtrx_qty_bkord As Decimal, _
                                           ByVal mpromise_dt As Integer, ByVal mrev_no As String, ByVal mdeall_amt As Decimal, _
                                           ByVal mfiller_0004 As String) As InventoryTrx

        Dim invTrx As New InventoryTrx

        invTrx.source = msource
        invTrx.ord_no = mord_no
        invTrx.ctl_no = mctl_no
        invTrx.line_no = mline_no
        invTrx.lev_no = mlev_no
        invTrx.seq_no = mseq_no
        invTrx.from_source = mfrom_source
        invTrx.from_ord_no = mfrom_ord_no
        invTrx.from_ctl_no = mfrom_ctl_no
        invTrx.from_line_no = mfrom_line_no
        invTrx.from_lev_no = mfrom_lev_no
        invTrx.from_seq_no = mfrom_seq_no
        invTrx.item_no = mitem_no
        invTrx.item_filler = mitem_filler
        invTrx.par_item_no = mpar_item_no
        invTrx.par_item_filler = mpar_item_filler
        invTrx.loc = mloc
        invTrx.trx_dt = mtrx_dt
        invTrx.trx_tm = mtrx_tm
        invTrx.doc_dt = mdoc_dt
        invTrx.doc_type = mdoc_type
        invTrx.doc_ord_no = mdoc_ord_no
        invTrx.doc_source = mdoc_source
        invTrx.cus_no = mcus_no
        invTrx.vend_no = mvend_no
        invTrx.prod_type = mprod_type
        invTrx.quantity = mquantity
        invTrx.old_quantity = mold_quantity
        invTrx.unit_cost = munit_cost
        invTrx.old_unit_cost = mold_unit_cost
        invTrx.new_unit_cost = mnew_unit_cost
        invTrx.price = mprice
        invTrx.build_qty = mbuild_qty
        invTrx.build_qty_per = mbuild_qty_per
        invTrx.amt = mamt
        invTrx.landed_cost = mlanded_cost
        invTrx.receipt_ord_no = mreceipt_ord_no
        invTrx.status = mstatus
        invTrx.jnl = mjnl
        invTrx.batch_id = mbatch_id
        invTrx.user_name = muser_name
        invTrx.id_no = mid_no
        invTrx.comment = mcomment
        invTrx.filler_0003 = mfiller_0003
        invTrx.trx_qty_bkord = mtrx_qty_bkord
        invTrx.promise_dt = mpromise_dt
        invTrx.rev_no = mrev_no
        invTrx.deall_amt = mdeall_amt
        invTrx.filler_0004 = mfiller_0004

        Return invTrx
    End Function

    Public Shared Function NewInventoryTrx() As InventoryTrx

        Dim invTrx As New InventoryTrx

        invTrx.source = ""
        invTrx.ord_no = ""
        invTrx.ctl_no = 0
        invTrx.line_no = 0
        invTrx.lev_no = 0
        invTrx.seq_no = 0
        invTrx.from_source = ""
        invTrx.from_ord_no = ""
        invTrx.from_ctl_no = 0
        invTrx.from_line_no = 0
        invTrx.from_lev_no = 0
        invTrx.from_seq_no = 0
        invTrx.item_no = ""
        invTrx.item_filler = ""
        invTrx.par_item_no = ""
        invTrx.par_item_filler = ""
        invTrx.loc = ""
        invTrx.trx_dt = 0
        invTrx.trx_tm = 0
        invTrx.doc_dt = 0
        invTrx.doc_type = ""
        invTrx.doc_ord_no = ""
        invTrx.doc_source = ""
        invTrx.cus_no = ""
        invTrx.vend_no = ""
        invTrx.prod_type = ""
        invTrx.quantity = 0
        invTrx.old_quantity = 0
        invTrx.unit_cost = 0
        invTrx.old_unit_cost = 0
        invTrx.new_unit_cost = 0
        invTrx.price = 0
        invTrx.build_qty = 0
        invTrx.build_qty_per = 0
        invTrx.amt = 0
        invTrx.landed_cost = 0
        invTrx.receipt_ord_no = ""
        invTrx.status = ""
        invTrx.jnl = ""
        invTrx.batch_id = ""
        invTrx.user_name = ""
        invTrx.id_no = ""
        invTrx.comment = ""
        invTrx.filler_0003 = ""
        invTrx.trx_qty_bkord = 0
        invTrx.promise_dt = 0
        invTrx.rev_no = ""
        invTrx.deall_amt = 0
        invTrx.filler_0004 = ""

        Return invTrx

    End Function

    Public Function GetInventoryLocation(ByVal item_no As String, ByVal loc As String, ByVal cn As SqlConnection) As InventoryTrx

        Dim invtrx As New InventoryTrx
        Dim dt As New DataTable
        Dim sSQL As String = "select top 1 * from " & TABLE & " " & vbCrLf & _
                             " where item_no = '" & item_no & "' and loc = '" & loc & "' " & vbCrLf & _
                             " order by A4GLIdentity DESC"

        dt = DAC.ExecuteSQL_DataTable(sSQL, cn, TABLE)

        invtrx = PopulateInventoryTransaction(dt)

        Return invtrx

    End Function

    Private Function PopulateInventoryTransaction(ByVal dt As DataTable) As InventoryTrx
        Dim invtrx As New InventoryTrx
        Dim obj(48) As Object
        For Each rw As DataRow In dt.Rows

            obj(0) = ReplaceNullValue(rw("source").ToString.Trim, dt.Columns("source").DataType.ToString.Trim)
            obj(1) = ReplaceNullValue(rw("ord_no").ToString.Trim, dt.Columns("ord_no").DataType.ToString.Trim)
            obj(2) = ReplaceNullValue(rw("ctl_no").ToString.Trim, dt.Columns("ctl_no").DataType.ToString.Trim)
            obj(3) = ReplaceNullValue(rw("line_no").ToString.Trim, dt.Columns("line_no").DataType.ToString.Trim)
            obj(4) = ReplaceNullValue(rw("lev_no").ToString.Trim, dt.Columns("lev_no").DataType.ToString.Trim)
            obj(5) = ReplaceNullValue(rw("seq_no").ToString.Trim, dt.Columns("seq_no").DataType.ToString.Trim)
            obj(6) = ReplaceNullValue(rw("from_source").ToString.Trim, dt.Columns("from_source").DataType.ToString.Trim)
            obj(7) = ReplaceNullValue(rw("from_ord_no").ToString.Trim, dt.Columns("from_ord_no").DataType.ToString.Trim)
            obj(8) = ReplaceNullValue(rw("from_ctl_no").ToString.Trim, dt.Columns("from_ctl_no").DataType.ToString.Trim)
            obj(9) = ReplaceNullValue(rw("from_line_no").ToString.Trim, dt.Columns("from_line_no").DataType.ToString.Trim)
            obj(10) = ReplaceNullValue(rw("from_lev_no").ToString.Trim, dt.Columns("from_lev_no").DataType.ToString.Trim)
            obj(11) = ReplaceNullValue(rw("from_seq_no").ToString.Trim, dt.Columns("from_seq_no").DataType.ToString.Trim)
            obj(12) = ReplaceNullValue(rw("item_no").ToString.Trim, dt.Columns("item_no").DataType.ToString.Trim)
            obj(13) = ReplaceNullValue(rw("item_filler").ToString.Trim, dt.Columns("item_filler").DataType.ToString.Trim)
            obj(14) = ReplaceNullValue(rw("par_item_no").ToString.Trim, dt.Columns("par_item_no").DataType.ToString.Trim)
            obj(15) = ReplaceNullValue(rw("par_item_filler").ToString.Trim, dt.Columns("par_item_filler").DataType.ToString.Trim)
            obj(16) = ReplaceNullValue(rw("loc").ToString.Trim, dt.Columns("loc").DataType.ToString.Trim)
            obj(17) = ReplaceNullValue(rw("trx_dt").ToString.Trim, dt.Columns("trx_dt").DataType.ToString.Trim)
            obj(18) = ReplaceNullValue(rw("trx_tm").ToString.Trim, dt.Columns("trx_tm").DataType.ToString.Trim)
            obj(19) = ReplaceNullValue(rw("doc_dt").ToString.Trim, dt.Columns("doc_dt").DataType.ToString.Trim)
            obj(20) = ReplaceNullValue(rw("doc_type").ToString.Trim, dt.Columns("doc_type").DataType.ToString.Trim)
            obj(21) = ReplaceNullValue(rw("doc_ord_no").ToString.Trim, dt.Columns("doc_ord_no").DataType.ToString.Trim)
            obj(22) = ReplaceNullValue(rw("doc_source").ToString.Trim, dt.Columns("doc_source").DataType.ToString.Trim)
            obj(23) = ReplaceNullValue(rw("cus_no").ToString.Trim, dt.Columns("cus_no").DataType.ToString.Trim)
            obj(24) = ReplaceNullValue(rw("vend_no").ToString.Trim, dt.Columns("vend_no").DataType.ToString.Trim)
            obj(25) = ReplaceNullValue(rw("prod_type").ToString.Trim, dt.Columns("prod_type").DataType.ToString.Trim)
            obj(26) = ReplaceNullValue(rw("quantity").ToString.Trim, dt.Columns("quantity").DataType.ToString.Trim)
            obj(27) = ReplaceNullValue(rw("old_quantity").ToString.Trim, dt.Columns("old_quantity").DataType.ToString.Trim)
            obj(28) = ReplaceNullValue(rw("unit_cost").ToString.Trim, dt.Columns("unit_cost").DataType.ToString.Trim)
            obj(29) = ReplaceNullValue(rw("old_unit_cost").ToString.Trim, dt.Columns("old_unit_cost").DataType.ToString.Trim)
            obj(30) = ReplaceNullValue(rw("new_unit_cost").ToString.Trim, dt.Columns("new_unit_cost").DataType.ToString.Trim)
            obj(31) = ReplaceNullValue(rw("price").ToString.Trim, dt.Columns("price").DataType.ToString.Trim)
            obj(32) = ReplaceNullValue(rw("build_qty").ToString.Trim, dt.Columns("build_qty").DataType.ToString.Trim)
            obj(33) = ReplaceNullValue(rw("build_qty_per").ToString.Trim, dt.Columns("build_qty_per").DataType.ToString.Trim)
            obj(34) = ReplaceNullValue(rw("amt").ToString.Trim, dt.Columns("amt").DataType.ToString.Trim)
            obj(35) = ReplaceNullValue(rw("landed_cost").ToString.Trim, dt.Columns("landed_cost").DataType.ToString.Trim)
            obj(36) = ReplaceNullValue(rw("receipt_ord_no").ToString.Trim, dt.Columns("receipt_ord_no").DataType.ToString.Trim)
            obj(37) = ReplaceNullValue(rw("status").ToString.Trim, dt.Columns("status").DataType.ToString.Trim)
            obj(38) = ReplaceNullValue(rw("jnl").ToString.Trim, dt.Columns("jnl").DataType.ToString.Trim)
            obj(39) = ReplaceNullValue(rw("batch_id").ToString.Trim, dt.Columns("batch_id").DataType.ToString.Trim)
            obj(40) = ReplaceNullValue(rw("user_name").ToString.Trim, dt.Columns("user_name").DataType.ToString.Trim)
            obj(41) = ReplaceNullValue(rw("id_no").ToString.Trim, dt.Columns("id_no").DataType.ToString.Trim)
            obj(42) = ReplaceNullValue(rw("comment").ToString.Trim, dt.Columns("comment").DataType.ToString.Trim)
            obj(43) = ReplaceNullValue(rw("filler_0003").ToString.Trim, dt.Columns("filler_0003").DataType.ToString.Trim)
            obj(44) = ReplaceNullValue(rw("trx_qty_bkord").ToString.Trim, dt.Columns("trx_qty_bkord").DataType.ToString.Trim)
            obj(45) = ReplaceNullValue(rw("promise_dt").ToString.Trim, dt.Columns("promise_dt").DataType.ToString.Trim)
            obj(46) = ReplaceNullValue(rw("rev_no").ToString.Trim, dt.Columns("rev_no").DataType.ToString.Trim)
            obj(47) = ReplaceNullValue(rw("deall_amt").ToString.Trim, dt.Columns("deall_amt").DataType.ToString.Trim)
            obj(48) = ReplaceNullValue(rw("filler_0004").ToString.Trim, dt.Columns("filler_0004").DataType.ToString.Trim)


        Next

        invtrx = NewInventoryTrx(
        obj(0).ToString.Trim,
        obj(1).ToString.Trim,
        Convert.ToInt32(obj(2)),
        Convert.ToInt32(obj(3)),
        Convert.ToInt32(obj(4)),
        Convert.ToInt32(obj(5)),
        obj(6).ToString.Trim,
        obj(7).ToString.Trim,
        Convert.ToInt32(obj(8)),
        Convert.ToInt32(obj(9)),
        Convert.ToInt32(obj(10)),
        Convert.ToInt32(obj(11)),
        obj(12).ToString.Trim,
        obj(13).ToString.Trim,
        obj(14).ToString.Trim,
        obj(15).ToString.Trim,
        obj(16).ToString.Trim,
        Convert.ToInt32(obj(17)),
        Convert.ToInt32(obj(18)),
        Convert.ToInt32(obj(19)),
        obj(20).ToString.Trim,
        obj(21).ToString.Trim,
        obj(22).ToString.Trim,
        obj(23).ToString.Trim,
        obj(24).ToString.Trim,
        obj(25).ToString.Trim,
        Convert.ToDecimal(obj(26)),
        Convert.ToDecimal(obj(27)),
        Convert.ToDecimal(obj(28)),
        Convert.ToDecimal(obj(29)),
        Convert.ToDecimal(obj(30)),
        Convert.ToDecimal(obj(31)),
        Convert.ToDecimal(obj(32)),
        Convert.ToDecimal(obj(33)),
        Convert.ToDecimal(obj(34)),
        Convert.ToDecimal(obj(35)),
        obj(36).ToString.Trim,
        obj(37).ToString.Trim,
        obj(38).ToString.Trim,
        obj(39).ToString.Trim,
        obj(40).ToString.Trim,
        obj(41).ToString.Trim,
        obj(42).ToString.Trim,
        obj(43).ToString.Trim,
        Convert.ToDecimal(obj(44)),
        Convert.ToInt32(obj(45)),
        obj(46).ToString.Trim,
        Convert.ToDecimal(obj(47)),
        obj(48).ToString.Trim
        )

        Return invtrx

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
