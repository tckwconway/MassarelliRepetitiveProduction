Imports System.Data.SqlClient

Public Class cInventoryBinTrx

    Public source As String
    Public ord_no As String
    Public ctl_no As Integer
    Public line_no As Integer
    Public lev_no As Integer
    Public seq_no As String
    Public bin_no As String
    Public quantity As Decimal
    Public unit_cost As Decimal
    Public alloc_dt As Integer
    Public trx_dt As Integer
    Public trx_tm As Integer
    Public item_no As String
    Public item_filler As String
    Public trx_qty_select As Decimal
    Public trx_eff_dt As Integer
    Public filler_0002 As String

    Const TABLE As String = "IMBINTRX_SQL"

    Public Shared Function NewInventoryBinTrx(ByVal m_source As String, ByVal m_ord_no As String, _
                                              ByVal m_ctl_no As Integer, ByVal m_line_no As Integer, ByVal m_lev_no As Integer, _
                                              ByVal m_seq_no As String, ByVal m_bin_no As String, ByVal m_quantity As Decimal, _
                                              ByVal m_unit_cost As Decimal, ByVal m_alloc_dt As Integer, ByVal m_trx_dt As Integer, _
                                              ByVal m_trx_tm As Integer, ByVal m_item_no As String, ByVal m_item_filler As String, _
                                              ByVal m_trx_qty_select As Decimal, ByVal m_trx_eff_dt As Integer, ByVal m_filler_0002 As String) As cInventoryBinTrx
        Dim binTrx As New cInventoryBinTrx
        binTrx.source = m_source
        binTrx.ord_no = m_ord_no
        binTrx.ctl_no = m_ctl_no
        binTrx.line_no = m_line_no
        binTrx.lev_no = m_lev_no
        binTrx.seq_no = m_seq_no
        binTrx.bin_no = m_bin_no
        binTrx.quantity = m_quantity
        binTrx.unit_cost = m_unit_cost
        binTrx.alloc_dt = m_alloc_dt
        binTrx.trx_dt = m_trx_dt
        binTrx.trx_tm = m_trx_tm
        binTrx.item_no = m_item_no
        binTrx.item_filler = m_item_filler
        binTrx.trx_qty_select = m_trx_qty_select
        binTrx.trx_eff_dt = m_trx_eff_dt
        binTrx.filler_0002 = m_filler_0002

        Return binTrx
    End Function

    Public Shared Function NewInventoryBinTrx(tranData As Object, cn As SqlConnection) As cInventoryBinTrx
        Dim sSQL As String
        Dim binTrx As New cInventoryBinTrx

        binTrx.source = "I"
        binTrx.ord_no = tranData.OrderNo.ToString.Trim
        binTrx.ctl_no = 0
        binTrx.line_no = 0
        binTrx.lev_no = tranData.LevelNo            '-- 0 for the first record, 1 for the second.  First record is the 'from' bin, second is the 'to' bin.  
        binTrx.seq_no = 0
        binTrx.bin_no = tranData.BinNo.ToString.Trim
        binTrx.quantity = tranData.qtyToMove
        binTrx.unit_cost = 0
        binTrx.alloc_dt = IIf(tranData.TranType = "Receipt", tranData.TransEffectiveDate, 0) 'Allocation Date applies to Receipt not Issue
        binTrx.trx_dt = cBusObj.GetMacolaDate(Now)
        binTrx.trx_tm = cBusObj.GetMacolaTime(Now)  'Time as integer 75428 as hhmmss= 7 hh, 54 mm, 28 ss, note no leading 0 because it's an integer
        binTrx.item_no = tranData.ItemNo.ToString.Trim
        binTrx.item_filler = ""
        binTrx.trx_qty_select = 0
        'HOW trx_eff_dt (Transaction Effective Date) IS CALCULATED - The mystery revealed! 
        ' Transfer: if lev_no = 0 (meaning it is issued from) then the trx_eff_dt is copied from the item's previous and most recent trx_eff_dt
        '           if lev_no = 1 (menaing it is received) then trx_eff_dt is today's date
        '           NOTE: lev_no = 1 is the only exception, Transfer and 0, Issue or Receipt all use today's date for the trx_eff_dt value
        ' Issue or Receipt Transactions
        '           trx_eff_dt is today's date
        If (tranData.TranType.ToString = "Transfer" And tranData.LevelNo.ToString = "0") OrElse tranData.trantype.ToString = "Issue" Then
            sSQL = "select IsNull(max(received_dt), 0) received_dt from IMINVBIN_SQL where item_no = '" & tranData.ItemNo & "' and bin_no = '" & tranData.BinNo & "'"
            binTrx.trx_eff_dt = Convert.ToInt32(DAC.Execute_Scalar(sSQL, cn))
        Else
            binTrx.trx_eff_dt = tranData.TransEffectiveDate
        End If
        binTrx.filler_0002 = ""

        Return binTrx

    End Function

End Class
