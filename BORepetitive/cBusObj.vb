Imports System.Data.SqlClient


Public Class cBusObj
    Private busObj As cBusObj

    Public Shared Function ExecuteSQLDataTable(ByVal sql As String, ByVal tablename As String, ByVal cn As SqlConnection) As DataTable
        Dim dt As DataTable

        dt = DAC.ExecuteSQL_DataTable(sql, cn, tablename)
        Return dt

    End Function

    Public Shared Function GetTransactionHistory(itm As String, ByVal cn As SqlConnection) As DataTable
        Dim dt As DataTable

        dt = DAC.ExecuteSP_DataTable(My.Resources.SP_spIMInventoryTransHistory, cn, _
                                     DAC.Parameter(My.Resources.Param_iItemNo, itm, ParameterDirection.Input))
        Return dt

    End Function

    Public Shared Function ExecuteSQLScalar(sql As String, cn As SqlConnection) As Object

        Dim o As Object
        o = DAC.Execute_Scalar(sql, cn)
        Return o

    End Function

    Public Shared Sub Execute_NonSQL(ssql, cn)
        DAC.Execute_NonSQL(ssql, cn)
    End Sub

    Public Shared Function FillIMInvBinWrk_DataTable(ByVal ord_no As String, ByVal tran_type As String, ByVal cn As SqlConnection) As DataTable

        Dim dt As DataTable = DAC.ExecuteSP_DataTable(My.Resources.SP_spIMInvBin_WRK, cn, _
                                         DAC.Parameter(My.Resources.Param_iord_no, ord_no, ParameterDirection.Input), _
                                         DAC.Parameter(My.Resources.Param_iTran_type, tran_type, ParameterDirection.Input))

        Return dt

    End Function
    Public Shared Function GetNextDocNo(sSQL As String, cn As SqlConnection) As Integer
        Dim next_doc_no As Integer = Convert.ToInt32(DAC.Execute_Scalar(sSQL, cn))
        Return next_doc_no
    End Function

    Public Shared Function GetCurrentBins(ItemNo As String, ItemCount As Integer, Loc As String, cn As SqlConnection) As DataTable
        Dim dt As New DataTable

        dt = DAC.ExecuteSP_DataTable(My.Resources.SP_spIMGetCurrentBins, cn, _
             DAC.Parameter(My.Resources.Param_iItemNo, ItemNo, ParameterDirection.Input), _
             DAC.Parameter(My.Resources.Param_iItemCount, ItemCount, ParameterDirection.Input), _
             DAC.Parameter(My.Resources.Param_iLoc, Loc, ParameterDirection.Input))

        Return dt
    End Function

    Public Shared Function GetQtyOnHandFromBin(itemNo As String, loc As String, bin As String, cn As SqlConnection) As DataTable
        Dim dt As DataTable
        dt = DAC.ExecuteSP_DataTable(My.Resources.SP_spIMGetQtyOnHand_MAS, cn, _
                                         DAC.Parameter(My.Resources.Param_iItemNo, itemNo, ParameterDirection.Input), _
                                         DAC.Parameter(My.Resources.Param_iLoc, loc, ParameterDirection.Input), _
                                         DAC.Parameter(My.Resources.Param_iBinTo, bin, ParameterDirection.Input))
        Return dt
    End Function
    Public Shared Function GetItemLocation(ItemNo As String, cn As SqlConnection) As DataTable
        Dim dt As Object

        dt = DAC.ExecuteSP_DataTable(My.Resources.SP_spIMGetItemLocation_MAS, cn, _
             DAC.Parameter(My.Resources.Param_iItemNo, ItemNo, ParameterDirection.Input))

        Return dt
    End Function



    Public Shared Function PrepareInventoryBinSQLTransaction(trantype As String, transaction As Object, TransDirection As String, SQLCommandType As String, cn As SqlConnection, Optional tabledata As Object = Nothing) As String
        Dim cInvBin As New InventoryBin
        Dim busObj As New cBusObj
        Dim sSQL As String = ""
        ' Dim itranEffectiveDate As Integer = GetMacolaDate(tranEffectiveDate.Year.ToString, Format(tranEffectiveDate.Month.ToString.PadLeft(2, "0")), Format(tranEffectiveDate.Day.ToString.PadLeft(2, "0")))
        Select Case trantype

            Case "T" 'Transfer

                'If SQLCommandType = "UPDATE" Then
                sSQL = busObj.BuildMySQLString("IMINVBIN_SQL", transaction, tabledata, SQLCommandType)
                'Else
                'sSQL = busObj.BuildMySQLString("IMINVBIN_SQL", SQLCommandType, transaction, tabledata)
                'End If

            Case "I", "R" 'Issue or Receipt
                'If SQLCommandType = "UPDATE" Then
                '    sSQL = busObj.BuildMySQLString("IMINVBIN_SQL", transaction, tabledata, SQLCommandType)
                'Else
                sSQL = busObj.BuildMySQLString("IMINVBIN_SQL", transaction, tabledata, SQLCommandType)
                'End If

                'Case "R"
                '    If SQLCommandType = "UPDATE" Then
                '        sSQL = busObj.BuildMySQLString("IMINVBIN_SQL", transaction, tabledata, SQLCommandType)
                '    Else
                '        sSQL = busObj.BuildMySQLString("IMINVBIN_SQL", transaction, tabledata, SQLCommandType)
                '    End If
        End Select

        Return sSQL

    End Function

    Public Shared Function PrepareBinTrxSQL(TranData As Object, cn As SqlConnection) As String
        Dim binTrx As New cInventoryBinTrx
        Dim busObj As New cBusObj
        Dim sSQL As String = ""

        binTrx = cInventoryBinTrx.NewInventoryBinTrx(TranData, cn)
        sSQL = busObj.BuildMySQLString("IMBINTRX_SQL", binTrx)

        Return sSQL

    End Function

    Public Shared Function PrepareInvLocSQL(TranData As Object, SQLCommandType As String, cn As SqlConnection, Optional TableData As Object = Nothing) As String

        Dim busObj As New cBusObj
        Dim sSQL As String = ""

        sSQL = busObj.BuildMySQLString("IMINVLOC_SQL", TranData, TableData, SQLCommandType)

        Return sSQL

    End Function

    Public Shared Function PrepareInvTrxSQL(TranData As Object, TableData As Object, cn As SqlConnection) As String

        Dim busObj As New cBusObj
        Dim sSQL As String = ""

        Dim invTrx = InventoryTrx.NewInventoryTrx(TranData, TableData, cn)
        sSQL = busObj.BuildMySQLString("IMINVTRX_SQL", invTrx)

        Return sSQL

    End Function

    Friend Function BuildMySQLString(tableName As String, ObjectData As Object, _
                                     Optional TableData As Object = Nothing, Optional SQLTranType As String = "") As String
        Dim sSQL As String = ""

        Select Case tableName
            Case "IMINVBIN_SQL"
                Select Case SQLTranType
                    Case "INSERT"

                        sSQL = "Insert Into IMINVBIN_SQL " & vbCrLf & _
                               "(item_no, item_filler, loc, bin_no, " & vbCrLf & _
                               " bin_priority, issue_priority, issue_pri_tm, bin_status, " & vbCrLf & _
                               " prev_status, received_dt, contract_no, unit_cost, " & vbCrLf & _
                               " orig_cost, qty_on_hand, qty_allocated, qty_bkord, " & vbCrLf & _
                               " qty_on_ord, cycle_count_cd, last_count_dt, tms_cntd_ytd, " & vbCrLf & _
                               " pct_err_lst_cnt, frz_cost, frz_qty, frz_dt, " & vbCrLf & _
                               " frz_tm, max_designator, max_value, max_uom, " & vbCrLf & _
                               " max_batch_post, cube_width_uom, cube_length_uom, cube_height_uom, " & vbCrLf & _
                               " cube_width, cube_length, cube_height, cube_qty_per, " & vbCrLf & _
                               " user_def_fld_1, user_def_fld_2, user_def_fld_3, user_def_fld_4, " & vbCrLf & _
                               " user_def_fld_5, tag_qty, filler_0003" & vbCrLf & _
                               ") " & vbCrLf & _
                               " Values " & vbCrLf & _
                               "('" & ObjectData.ItemNo.ToString.Trim & "',  '', '" & ObjectData.Loc.ToString.Trim & "', '" & ObjectData.BinNo.ToString & "', " & vbCrLf & _
                               " " & ObjectData.BinPriority & ", '" & TableData.issue_priority.ToString.Trim & "', 0, '" & ObjectData.BinStatus & "', " & vbCrLf & _
                               " NULL, " & ObjectData.ReceivedDt & ", NULL, 0, " & vbCrLf & _
                               " 0, (" & ObjectData.QtyOnHand & ObjectData.CalculationOperator & ObjectData.QtyToMove & "), 0, 0, " & vbCrLf & _
                               " 0, NULL, 0, 0," & vbCrLf & _
                               " 0, 0, 0, 0, " & vbCrLf & _
                               " 0, NULL, 0, NULL, " & vbCrLf & _
                               " 0, NULL, NULL, NULL, " & vbCrLf & _
                               " 0, 0, 0, 0, " & vbCrLf & _
                               " NULL, NULL, NULL, NULL, " & vbCrLf & _
                               " NULL, 0 ,NULL " & vbCrLf & _
                               ") "
                    Case "UPDATE"
                        'add existing qty to UPDATE qty

                        sSQL = "Update IMINVBIN_SQL  " & vbCrLf &
                                                      "Set " & vbCrLf & _
                                                      " qty_on_hand = (" & ObjectData.QtyOnHand & ObjectData.CalculationOperator & ObjectData.QtyToMove & ") " & vbCrLf & _
                                                      ",received_dt = " & ObjectData.ReceivedDt & " " & vbCrLf & _
                                                      ",bin_priority = " & ObjectData.BinPriority & " " & vbCrLf & _
                                                      ",bin_status = '" & ObjectData.BinStatus & "' " & vbCrLf & _
                                                      "Where item_no = '" & ObjectData.ItemNo.ToString.Trim & "' " & vbCrLf & _
                                                      "   and bin_no = '" & ObjectData.BinNo.ToString.Trim & "' " & vbCrLf & _
                                                      "   and loc = '" & ObjectData.Loc.ToString.Trim & "' "

                End Select

            Case "IMBINTRX_SQL"
                Dim bintrx As cInventoryBinTrx = CType(ObjectData, cInventoryBinTrx)
                With bintrx

                    sSQL = "Insert Into IMBINTRX_SQL " & vbCrLf & _
                           "(" & vbCrLf & _
                           "source, ord_no, ctl_no, line_no, lev_no, seq_no, bin_no, quantity, unit_cost, alloc_dt, " & vbCrLf & _
                           "trx_dt, trx_tm, item_no, item_filler, trx_qty_select, trx_eff_dt, filler_0002" & vbCrLf & _
                           ")" & vbCrLf & _
                           "Values " & vbCrLf & _
                           "( " & vbCrLf & _
                           " '" & .source & "' " & vbCrLf & _
                           ",'" & .ord_no & "' " & vbCrLf & _
                           "," & .ctl_no & " " & vbCrLf & _
                           "," & .line_no & " " & vbCrLf & _
                           "," & .lev_no & " " & vbCrLf & _
                           "," & .seq_no & " " & vbCrLf & _
                           ",'" & .bin_no & "' " & vbCrLf & _
                           "," & .quantity & " " & vbCrLf & _
                           "," & .unit_cost & " " & vbCrLf & _
                           "," & .alloc_dt & " " & vbCrLf & _
                           "," & .trx_dt & " " & vbCrLf & _
                           "," & .trx_tm & " " & vbCrLf & _
                           ",'" & .item_no & "' " & vbCrLf & _
                           ",'" & .item_filler & "' " & vbCrLf & _
                           "," & .trx_qty_select & " " & vbCrLf & _
                           "," & .trx_eff_dt & " " & vbCrLf & _
                           ",'" & .filler_0002 & "' " & vbCrLf & _
                           ") "

                End With

            Case "IMINVLOC_SQL"
                'Set the operator to + or - depending if we're moving stock in or out.  
                'NOTE: UPDATES ONLY on INIMVLOB  - in Macola Inventory Transactions are only possible if the item already exists in Item Master, 
                '      on the fly' adding of items is not allowed.  So we don't need an INSERT version for IMINVLOC_SQL Table.  

                'Dim sOperator As String = IIf(Convert.ToInt32(ObjectData.levelno) = 0, " - ", " + ")

                sSQL = "Update IMINVLOC_SQL  " & vbCrLf & _
                              "   Set qty_on_hand = (" & ObjectData.QtyOnHand & ObjectData.CalculationOperator & ObjectData.QtyToMove & "), " & vbCrLf & _
                              "         usage_ptd = " & ObjectData.usageptd & ", " & vbCrLf & _
                              "         usage_ytd = " & ObjectData.usageytd & " " & vbCrLf & _
                              " Where item_no = '" & ObjectData.ItemNo.ToString.Trim & "'" & vbCrLf & _
                              "   and loc = '" & ObjectData.Loc.ToString.Trim & "'"

            Case "IMINVTRX_SQL"

                Dim invtrx As InventoryTrx = CType(ObjectData, InventoryTrx)
                With invtrx
                    sSQL = "Insert Into IMINVTRX_SQL " & vbCrLf & _
                    "(" & vbCrLf & _
                           " source, ord_no, ctl_no, line_no, lev_no, seq_no, from_source, from_ord_no, from_ctl_no,  " & vbCrLf & _
                           " from_line_no, from_lev_no, from_seq_no, item_no, item_filler, par_item_no, par_item_filler,  " & vbCrLf & _
                           " loc, trx_dt, trx_tm, doc_dt, doc_type, doc_ord_no, doc_source, cus_no, vend_no, prod_type,  " & vbCrLf & _
                           " quantity, old_quantity, unit_cost, old_unit_cost, new_unit_cost, price, build_qty, " & vbCrLf & _
                           " build_qty_per, amt, landed_cost, receipt_ord_no, status, jnl, batch_id, user_name, id_no,  " & vbCrLf & _
                           " comment, filler_0003, trx_qty_bkord, promise_dt, rev_no, deall_amt, filler_0004 " & vbCrLf & _
                           ")" & vbCrLf & _
                           "Values " & vbCrLf & _
                           "( " & vbCrLf & _
                           " '" & .source & "' " & vbCrLf & _
                           ",'" & .ord_no & "' " & vbCrLf & _
                           "," & .ctl_no & " " & vbCrLf & _
                           "," & .line_no & " " & vbCrLf & _
                           "," & .lev_no & " " & vbCrLf & _
                           "," & .seq_no & " " & vbCrLf & _
                           ",'" & .from_source & "' " & vbCrLf & _
                           ",'" & .from_ord_no & "' " & vbCrLf & _
                           "," & .from_ctl_no & " " & vbCrLf & _
                           "," & .from_line_no & " " & vbCrLf & _
                           "," & .from_lev_no & " " & vbCrLf & _
                           "," & .from_seq_no & " " & vbCrLf & _
                           ",'" & .item_no & "' " & vbCrLf & _
                           ",'" & .item_filler & "' " & vbCrLf & _
                           ",'" & .par_item_no & "' " & vbCrLf & _
                           ",'" & .par_item_filler & "' " & vbCrLf & _
                           ",'" & .loc & "' " & vbCrLf & _
                           "," & .trx_dt & " " & vbCrLf & _
                           "," & .trx_tm & " " & vbCrLf & _
                           "," & .doc_dt & " " & vbCrLf & _
                           ",'" & .doc_type & "' " & vbCrLf & _
                           ",'" & .doc_ord_no & "' " & vbCrLf & _
                           ",'" & .doc_source & "' " & vbCrLf & _
                           ",'" & .cus_no & "' " & vbCrLf & _
                           ",'" & .vend_no & "' " & vbCrLf & _
                           ",'" & .prod_type & "' " & vbCrLf & _
                           "," & .quantity & " " & vbCrLf & _
                           "," & .old_quantity & " " & vbCrLf & _
                           "," & .unit_cost & " " & vbCrLf & _
                           "," & .old_unit_cost & " " & vbCrLf & _
                           "," & .new_unit_cost & " " & vbCrLf & _
                           "," & .price & " " & vbCrLf & _
                           "," & .build_qty & " " & vbCrLf & _
                           "," & .build_qty_per & " " & vbCrLf & _
                           "," & .amt & " " & vbCrLf & _
                           "," & .landed_cost & " " & vbCrLf & _
                           ",'" & .receipt_ord_no & "' " & vbCrLf & _
                           ",'" & .status & "' " & vbCrLf & _
                           ",'" & .jnl & "' " & vbCrLf & _
                           ",'" & .batch_id & "' " & vbCrLf & _
                           ",'" & .user_name & "' " & vbCrLf & _
                           ",'" & .id_no & "' " & vbCrLf & _
                           ",'" & .comment & "' " & vbCrLf & _
                           ",'" & .filler_0003 & "' " & vbCrLf & _
                           "," & .trx_qty_bkord & " " & vbCrLf & _
                           "," & .promise_dt & " " & vbCrLf & _
                           ",'" & .rev_no & "' " & vbCrLf & _
                           "," & .deall_amt & " " & vbCrLf & _
                           ",'" & .filler_0004 & "' " & vbCrLf & _
                           ") "
                End With

        End Select

        Return sSQL
    End Function


    Private Function GetSQLString(TableName As String, SQLTranType As String) As String
        Dim sSQL As String = ""

        Select Case TableName
            Case "IMINVBIN_SQL"

                If SQLTranType = "INSERT" Then
                    sSQL = "insert into IMINVBIN_SQL ( " & _
                            " item_no, item_filler, loc, bin_no, bin_priority, issue_priority, issue_pri_tm, bin_status, prev_status, received_dt, " & _
                            " contract_no, unit_cost, orig_cost, qty_on_hand, qty_allocated, qty_bkord, qty_on_ord, cycle_count_cd, last_count_dt, " & _
                            " tms_cntd_ytd, pct_err_lst_cnt, frz_cost, frz_qty, frz_dt, frz_tm, max_designator, max_value, max_uom, max_batch_post, " & _
                            " cube_width_uom, cube_length_uom, cube_height_uom, cube_width, cube_length, cube_height, cube_qty_per, " & _
                            " user_def_fld_1, user_def_fld_2, user_def_fld_3, user_def_fld_4, user_def_fld_5, tag_qty, filler_0003 " & _
                            ")"

                End If
        End Select
        Return sSQL

    End Function

    Public Shared Sub SaveRepetitiveLog(TranType As String, ItenNo As String, LocFrom As String, LocTo As String, _
                                        LocFromQtyOnHand As Integer, LocToQtyOnHand As Integer, BinFrom As String, BinTo As String, _
                                        BinFromQtyOnHand As Integer, BinToQtyOnHand As Integer, BinFromTrxQty As Integer, BinToTrxQty As Integer, _
                                        QtyToMove As Integer, CreateDate As Integer, OrderNo As String, TransactionID As Integer, SessionID As Integer, _
                                        cn As SqlConnection)

        DAC.ExecuteSP(My.Resources.SP_spIMSaveRepetitiveLog_MAS, cn, _
                        DAC.Parameter(My.Resources.Param_iTranType, TranType, ParameterDirection.Input), _
                        DAC.Parameter(My.Resources.Param_iItemNo, ItenNo, ParameterDirection.Input), _
                        DAC.Parameter(My.Resources.Param_iLocFrom, LocFrom, ParameterDirection.Input), _
                        DAC.Parameter(My.Resources.Param_iLocTo, LocTo, ParameterDirection.Input), _
                        DAC.Parameter(My.Resources.Param_iLocFromQtyOnHand, LocFromQtyOnHand, ParameterDirection.Input), _
                        DAC.Parameter(My.Resources.Param_iLocToQtyOnHand, LocToQtyOnHand, ParameterDirection.Input), _
                        DAC.Parameter(My.Resources.Param_iBinFrom, BinFrom, ParameterDirection.Input), _
                        DAC.Parameter(My.Resources.Param_iBinTo, BinTo, ParameterDirection.Input), _
                        DAC.Parameter(My.Resources.Param_iBinFromQtyOnHand, BinFromQtyOnHand, ParameterDirection.Input), _
                        DAC.Parameter(My.Resources.Param_iBinToQtyOnHand, BinToQtyOnHand, ParameterDirection.Input), _
                        DAC.Parameter(My.Resources.Param_iBinFromTrxQty, BinFromTrxQty, ParameterDirection.Input), _
                        DAC.Parameter(My.Resources.Param_iBinToTrxQty, BinToTrxQty, ParameterDirection.Input), _
                        DAC.Parameter(My.Resources.Param_iQtyToMove, QtyToMove, ParameterDirection.Input), _
                        DAC.Parameter(My.Resources.Param_iCreateDate, CreateDate, ParameterDirection.Input), _
                        DAC.Parameter(My.Resources.Param_iOrderNo, OrderNo, ParameterDirection.Input), _
                        DAC.Parameter(My.Resources.Param_iTransactionID, TransactionID, ParameterDirection.Input), _
                        DAC.Parameter(My.Resources.Param_iSessionID, SessionID, ParameterDirection.Input))

    End Sub
    Public Shared Function GetMacolaDate(Year As String, Month As String, Day As String) As Integer
        
        Return CInt(Year & Month & Day)

    End Function
    Public Shared Function GetMacolaDate(dt As Date) As Integer
        Dim yr As String = dt.Year.ToString
        Dim mo As String = Format(dt.Month.ToString.PadLeft(2, "0"))
        Dim dy As String = Format(dt.Day.ToString.PadLeft(2, "0"))

        Return CInt(yr & mo & dy)

    End Function

    'Public Shared Function GetMacolaTime(Hour As String, Minute As String, Second As String) As Integer
    '    Return CInt(Hour & Minute & Second)
    'End Function

    Public Shared Function GetMacolaTime(dt As Date) As Integer
        Dim hr As String = Format(dt.Hour.ToString.PadLeft(2, "0"))
        Dim mn As String = Format(dt.Minute.ToString.PadLeft(2, "0"))
        Dim sc As String = Format(dt.Second.ToString.PadLeft(2, "0"))
       
        Return CInt(hr & mn & sc)
    End Function

    Public Shared Function ValidateTransactionDate(trandate As Date, currPeriod As Integer, currYear As Integer)
        
        If currYear = trandate.Year AndAlso trandate.Month = currPeriod Then
            Return True
        Else
            Return False
        End If

    End Function
End Class
