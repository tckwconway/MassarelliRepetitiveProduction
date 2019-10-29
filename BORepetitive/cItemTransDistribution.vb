Imports System.ComponentModel
Imports System.Data.DataRowView
Imports System.Data.SqlClient
Imports System.Data
Imports System.Text
Imports System.Xml
Imports System.IO
Imports System.Data.SqlTypes

Public Class ItemTransDistribution


    Public item_no As String
    Public item_filler As String
    Public doc_no As String
    Public doc_dt As Integer
    Public seq_no As Integer
    Public dist_doc As Integer
    Public dist_sb As Integer
    Public mn_no As String
    Public sb_no As String
    Public dp_no As String
    Public pkg_id As String
    Public jnl_src As String
    Public job_no As String
    Public offset_mn_no As String
    Public offset_sb_no As String
    Public offset_dp_no As String
    Public dist_amt As Decimal
    Public dist_qty As Decimal
    Public reference As String
    Public filler_0002 As String

    Const TABLE As String = "IMTRXDST_SQL"


    Private Function NewItemTransDistribution(ByVal m_item_no As String, ByVal m_item_filler As String, ByVal m_doc_no As String, ByVal m_doc_dt As Integer, ByVal m_seq_no As Integer, _
                                                     ByVal m_dist_doc As Integer, ByVal m_dist_sb As Integer, ByVal m_mn_no As String, ByVal m_sb_no As String, ByVal m_dp_no As String, _
                                                     ByVal m_pkg_id As String, ByVal m_jnl_src As String, ByVal m_job_no As String, ByVal m_offset_mn_no As String, ByVal m_offset_sb_no As String, _
                                                     ByVal m_offset_dp_no As String, ByVal m_dist_amt As Decimal, ByVal m_dist_qty As Decimal, ByVal m_reference As String, ByVal m_filler_0002 As String) As ItemTransDistribution

        Dim imtrxdst As New ItemTransDistribution
        imtrxdst.item_no = m_item_no
        imtrxdst.item_filler = m_item_filler
        imtrxdst.doc_no = m_doc_no
        imtrxdst.doc_dt = m_doc_dt
        imtrxdst.seq_no = m_seq_no
        imtrxdst.dist_doc = m_dist_doc
        imtrxdst.dist_sb = m_dist_sb
        imtrxdst.mn_no = m_mn_no
        imtrxdst.sb_no = m_sb_no
        imtrxdst.dp_no = m_dp_no
        imtrxdst.pkg_id = m_pkg_id
        imtrxdst.jnl_src = m_jnl_src
        imtrxdst.job_no = m_job_no
        imtrxdst.offset_mn_no = m_offset_mn_no
        imtrxdst.offset_sb_no = m_offset_sb_no
        imtrxdst.offset_dp_no = m_offset_dp_no
        imtrxdst.dist_amt = m_dist_amt
        imtrxdst.dist_qty = m_dist_qty
        imtrxdst.reference = m_reference
        imtrxdst.filler_0002 = m_filler_0002

        Return imtrxdst

    End Function


    Public Function GetItemTransDistribution(ByVal item_no As String, ByVal doc_no As String, ByVal doc_dt As Integer, ByVal seq_no As Integer, ByVal cn As SqlConnection) As ItemTransDistribution

        Dim imtrxdst As New ItemTransDistribution
        Dim dt As New DataTable
        Dim sSQL As String = "Select * from " & TABLE & " " _
                           & "Where item_no = '" & item_no & "' and doc_no = '" & doc_no & "' and doc_dt = " & doc_dt & " and seq_no = " & seq_no & " " _
                           & " order by item_no, doc_no, doc_dt, seq_no "

        dt = DAC.ExecuteSQL_DataTable(sSQL, cn, TABLE)

        imtrxdst = PopulateItemTransDistribution(dt)

        Return imtrxdst

    End Function


    Private Function PopulateItemTransDistribution(ByVal dt As DataTable) As ItemTransDistribution
        Dim imtrxdst As New ItemTransDistribution
        Dim obj(19) As Object
        For Each rw As DataRow In dt.Rows

            obj(0) = ReplaceNullValue(rw("item_no"), dt.Columns("item_no").DataType.ToString)
            obj(1) = ReplaceNullValue(rw("item_filler"), dt.Columns("item_filler").DataType.ToString)
            obj(2) = ReplaceNullValue(rw("doc_no"), dt.Columns("doc_no").DataType.ToString)
            obj(3) = ReplaceNullValue(rw("doc_dt"), dt.Columns("doc_dt").DataType.ToString)
            obj(4) = ReplaceNullValue(rw("seq_no"), dt.Columns("seq_no").DataType.ToString)
            obj(5) = ReplaceNullValue(rw("dist_doc"), dt.Columns("dist_doc").DataType.ToString)
            obj(6) = ReplaceNullValue(rw("dist_sb"), dt.Columns("dist_sb").DataType.ToString)
            obj(7) = ReplaceNullValue(rw("mn_no"), dt.Columns("mn_no").DataType.ToString)
            obj(8) = ReplaceNullValue(rw("sb_no"), dt.Columns("sb_no").DataType.ToString)
            obj(9) = ReplaceNullValue(rw("dp_no"), dt.Columns("dp_no").DataType.ToString)
            obj(10) = ReplaceNullValue(rw("pkg_id"), dt.Columns("pkg_id").DataType.ToString)
            obj(11) = ReplaceNullValue(rw("jnl_src"), dt.Columns("jnl_src").DataType.ToString)
            obj(12) = ReplaceNullValue(rw("job_no"), dt.Columns("job_no").DataType.ToString)
            obj(13) = ReplaceNullValue(rw("offset_mn_no"), dt.Columns("offset_mn_no").DataType.ToString)
            obj(14) = ReplaceNullValue(rw("offset_sb_no"), dt.Columns("offset_sb_no").DataType.ToString)
            obj(15) = ReplaceNullValue(rw("offset_dp_no"), dt.Columns("offset_dp_no").DataType.ToString)
            obj(16) = ReplaceNullValue(rw("dist_amt"), dt.Columns("dist_amt").DataType.ToString)
            obj(17) = ReplaceNullValue(rw("dist_qty"), dt.Columns("dist_qty").DataType.ToString)
            obj(18) = ReplaceNullValue(rw("reference"), dt.Columns("reference").DataType.ToString)
            obj(19) = ReplaceNullValue(rw("filler_0002"), dt.Columns("filler_0002").DataType.ToString)



        Next

        imtrxdst = NewItemTransDistribution(
        obj(0).ToString,
        obj(1).ToString,
        obj(2).ToString,
        Convert.ToInt32(obj(3)),
        Convert.ToInt32(obj(4)),
        Convert.ToInt32(obj(5)),
        Convert.ToInt32(obj(6)),
        obj(7).ToString,
        obj(8).ToString,
        obj(9).ToString,
        obj(10).ToString,
        obj(11).ToString,
        obj(12).ToString,
        obj(13).ToString,
        obj(14).ToString,
        obj(15).ToString,
        Convert.ToDecimal(obj(16)),
        Convert.ToDecimal(obj(17)),
        obj(18).ToString,
        obj(19).ToString)

        Return imtrxdst

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

    Public Function InsertIntoIMTRXDST_SQL(ByVal m_item_no As String, ByVal m_item_filler As String, ByVal m_doc_no As String, ByVal m_doc_dt As Integer, ByVal m_seq_no As Integer, _
                                                     ByVal m_dist_doc As Integer, ByVal m_dist_sb As Integer, ByVal m_mn_no As String, ByVal m_sb_no As String, ByVal m_dp_no As String, _
                                                     ByVal m_pkg_id As String, ByVal m_jnl_src As String, ByVal m_job_no As String, ByVal m_offset_mn_no As String, ByVal m_offset_sb_no As String, _
                                                     ByVal m_offset_dp_no As String, ByVal m_dist_amt As Decimal, ByVal m_dist_qty As Decimal, ByVal m_reference As String, ByVal m_filler_0002 As String, _
                                                     ByVal cn As SqlConnection) As Boolean

        Dim sSQL As String = "INSERT INTO " & TABLE & " item_no,item_filler,doc_no,doc_dt,seq_no,dist_doc,dist_sb,mn_no,sb_no,dp_no,pkg_id,jnl_src,job_no,offset_mn_no,offset_sb_no,offset_dp_no,dist_amt,dist_qty,reference,filler_0002 " & _
                             "VALUES ('" & m_item_no & "', '" & m_item_filler & "', '" & m_doc_no & "', '" & m_doc_dt & "', '" & m_seq_no & "', '" & m_dist_doc & "', '" & m_dist_sb & "', '" & m_mn_no & "', '" & m_sb_no & "', '" _
                                         & m_dp_no & "', '" & m_pkg_id & "', '" & m_jnl_src & "', '" & m_job_no & "', '" & m_offset_mn_no & "', '" & m_offset_sb_no & "', '" & m_offset_dp_no & "', '" & m_dist_amt & "', '" _
                                         & m_dist_qty & "', '" & m_reference & "', '" & m_filler_0002 & "' )"
        DAC.Execute_NonSQL(sSQL, cn)

        Return Nothing

    End Function
End Class
