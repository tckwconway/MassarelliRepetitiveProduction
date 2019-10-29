Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Text

Public Class frmRepetitive
    Implements IDisposable

    'Classes
    Private company As CompanySetting
    Private imctlsetting As IMControlSettings
    Private invbin As InventoryBin
    Private invloc As InventoryLocation
    Private invtrx As InventoryTrx
    Private itmloc001 As ItemLocationSetting
    Private itmloc002 As ItemLocationSetting
    Private itmtrxdst As ItemTransDistribution
    Private sessiondata As New CurrentSessionData
    Private sysacct As SystemAcctSetting
    Private uidata As New UICurrentData

    Friend systemperiod As cSystemPeriodSettings

    'Collections
    Private cUIDataCollection As Collection

    'Lists
    Private lUIDataList As List(Of UICurrentData)
    Private SQLTransactionList As New List(Of String)

    'Data Tables
    Private dtAllFactoryBins As New DataTable
    Private dtAllItems As DataTable
    Private dtAllWarehouseBins As New DataTable
    Private dtFactory001 As New DataTable
    Private dtFactoryForCombo As New DataTable
    Private dtItem As New DataTable
    Private dtTransactionHistory As New DataTable
    Private dtTransferList As New DataTable
    Private dtWarehouse002 As New DataTable

    'Data Views
    Private dvFactory As DataView
    Private dvFactoryForCombo As DataView
    Private dvWareHouse As DataView

    'DataGridView ComboBoxColumn
    Private dgvComboBinFirst As New DataGridViewComboBoxColumn
    Private dgvComboBinSecond As New DataGridViewComboBoxColumn
    Private dgvComboLocFrom As New DataGridViewComboBoxColumn
    Private dgvComboLocTo As New DataGridViewComboBoxColumn

    'ComboBoxes Used in DataGridView
    Private WithEvents comboBinFirst As ComboBox
    Private WithEvents comboBinSecond As ComboBox
    Private WithEvents comboLocationTo As ComboBox
    Public WithEvents comboLocationFrom As ComboBox

    'Form Declaration (used on startup to retrieve the date from the fTransactionDate from)
    Private WithEvents fTransactionDate1 As fTransactionDate
    Private WithEvents fFactoryBinSelect1 As fFactoryBinSelect

    'DataGridView (used with DGV Edit Control Showing to determine which DGV we are on)
    Private dgvFromEditControlShowing As DataGridView

    'Private strip As New ContextMenuStrip()
    Private mouseLocation As DataGridViewCellEventArgs
    
    'Hit Test
    Private hitContextMenu As DataGridView.HitTestInfo
    Private ht As DataGridView.HitTestInfo

    'Booleans
    'Private bDeleteBin As Boolean = False
    Private bPaintDGVBorders As Boolean = False
    Private bAppLoading As Boolean = True
    Private bKeepFactoryBin As Boolean = True
    Private bSkipPaint As Boolean = False
    'DataGridView CellStyles (used with Transaction History DataGridView)
    Private styleAlternatingRowColor As DataGridViewCellStyle

    'Integers
    Private iDoneCounter As Integer = 0

    'TextBox
    Private WithEvents txtDGVQuantity As New TextBox

    'Transaction Date is set at strartup and is validated each time a transaction is started to be sure we have not diverged from
    'the current period for the IM Module.
    Public TransactionEffectiveDate As Date
    Public TransEffectiveDate As Integer
    Public TransferToBin As String

    'Colors
    Public ReadOnly clrTransfer As System.Drawing.Color = Color.FromArgb(255, 168, 195, 255)
    Public ReadOnly clrIssue As System.Drawing.Color = Color.FromArgb(255, 243, 232, 153)
    Public ReadOnly clrReceipt As System.Drawing.Color = Color.FromArgb(255, 173, 214, 88)

    Const emptyGL As String = "00000000"                      'This is the sb_no and db_no in table SYACTFIL_SQL, sub GL numbers, both "00000000" because sb_no and db_no values are not used for any GL Acct Nos at Massarelli...
    Const inventoryGL As String = "10400000"                  'This is the mn_no in table SYACTFIL_SQL
    Const sArrowTriangleLeft As Char = ChrW(&H25C0)           'Triangle Arrow Pointing Left
    Const sArrowTriangleLeftRight As Char = ChrW(&H2B0C)      'Triangle Arrow Pointing Left and Right
    Const sArrowTriangleRight As Char = ChrW(&H25B6)          'Triangle Arrow Pointing Left
    Const sCheckMark1 As Char = ChrW(&H2611)                  'Check with box
    Const sCheckMark2 As Char = ChrW(&H2713)                  'Light check mark
    Const sCheckMark3 As Char = ChrW(&H2714)                  'Heavy check mark
    Const sElipsis As Char = ChrW(&H2026)                     'Elipsis Horizontal
    Const sGlyphDown As Char = ChrW(&H25BC)                   'Glyph (down pointing triangle)
    Const sGlyphUp As Char = ChrW(&H25B2)                     'Glyph (up pointing triangle)
    Const sTransactionIN As String = "IN"
    Const sTransactionOUT As String = "OUT"
    Const sTransactionTRAN As String = "TRAN"
    Const sXLargeCrossMark As Char = ChrW(&H274C)             'Triangle Arrow Pointing Left
    Const sXSmallCrossMark As Char = ChrW(&H2A2F)             'Triangle Arrow Pointing Left

#Region "Enums"

    Private Enum Locations As Integer
        Warehouse = 2
        Factory = 1
    End Enum
    Private Enum TransferType As Integer
        Transfer = 1
        Issue = 2
        Receipt = 3
    End Enum
    Private Enum ReturnDataType As Integer
        Target = 1
        BinFrom = 2
        BinTo = 3
        Direction = 4
    End Enum

    Private Enum Bins As Integer
        BinFirst = 1
        BinSecond = 2
    End Enum

    Private Enum SQLCommandTypeEnum As Integer
        INSERT = 1
        UPDATE = 2
    End Enum

    Private Enum IM_CTL_KEY_1
        Massarelli = 1
        Unknown = 2
    End Enum

    Private Enum ValidateType As Integer
        ValidQuantity = 1
        Duplicated = 2
    End Enum

    Private Enum RepetitiveLogType As Integer
        SessionID = 1
        CreateDate = 2
    End Enum

#End Region

#Region "  REGION - Load Form and Startup  "

    Private Sub fRepetitive_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        MacStartup("DATA")
        'Load basic info on Macola IM to get started ...
        LoadInventoryCompanyControlFile()
        'AutoComplete is for the search for item numbers text box ...
        LoadAutoComplete()
        'Setup various tables used for comboboxes bins and items...
        LoadEmptyDataTables()
        'Setup the datagridviews ...
        LoadDataGridViews()
        'PictureBox Tags are populated with "Issue", "Receipt" or "Transfer" and used in the Drag/Drop to retrieve which PictureBox we have dropped on.  
        LoadPicBoxTags()
        'Get the IM Period, both Period and Date
        LoadCurrentPeriod()
        'Session IDs allow us to separate identical transactions.  
        LoadSessionData()
        'Load Data here populates the 
        LoadData(bAppLoading)

        With Timer1
            .Interval = 50
            .Enabled = True
        End With
    End Sub

#End Region

#Region "  REGION - Load Methods   "

    Private Sub LoadInventoryCompanyControlFile()
        sysacct = New SystemAcctSetting
        company = New CompanySetting
        imctlsetting = New IMControlSettings
        systemperiod = New cSystemPeriodSettings

        sysacct = sysacct.GetSystemAcctSetting(inventoryGL, emptyGL, emptyGL, cn)
        company = company.GetCompanySetting("1", cn)
        imctlsetting = imctlsetting.GetIMControlSetting(1, cn)
        systemperiod = systemperiod.GetSystemPeriodSetting(cSystemPeriodSettings.PeriodSetting.Inventory, cn)
        'itmloc001 = itmloc001.GetItemLocationSetting(itmloc001.PadWarehouseLocation(ItemLocationSetting.ItemLoc.Factory), cn)
        'itmloc002 = itmloc002.GetItemLocationSetting(itmloc002.PadWarehouseLocation(ItemLocationSetting.ItemLoc.Warehouse), cn)
        'invbin = invbin.GetInventoryBinSetting("M2156", itmloc001.PadWarehouseLocation(ItemLocationSetting.ItemLoc.Factory), cn)
        'itmtrxdst = itmtrxdst.GetItemTransDistribution("M2156", "", 0, 0, cn)

    End Sub

    Private Sub LoadAutoComplete()
        Dim sSQL As String
        Dim dt As New DataTable
        Dim ItmAutoCmplt As New AutoCompleteStringCollection

        sSQL = "Select distinct RTrim(lTrim(Substring(itm.item_no, 2, 15))) as item_no from IMITMIDX_SQL itm where left(item_no, 1) = 'M' "

        Try
            dt = cBusObj.ExecuteSQLDataTable(sSQL, "ItemsAutoComplete", cn)
        Catch ex As Exception
            MsgBox(ex.Message)

        End Try

        'Use Linq to get the Array ...
        Dim itms = (From row In dt Select itmno = row(0).ToString).ToArray

        ItmAutoCmplt.AddRange(itms)
        With txtItemSearch
            .AutoCompleteCustomSource = ItmAutoCmplt
            .AutoCompleteSource = AutoCompleteSource.CustomSource
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
        End With

    End Sub

    Private Sub LoadEmptyDataTables()

        With dtItem
            Dim item_no As DataColumn = New DataColumn("item_no", Type.GetType("System.String"))
            Dim item_desc_1 As DataColumn = New DataColumn("item_desc_1", Type.GetType("System.String"))

            .Columns.Add(item_no)
            .Columns.Add(item_desc_1)
        End With

        With dtFactory001
            Dim item_no As DataColumn = New DataColumn("item_no", Type.GetType("System.String"))
            Dim loc As DataColumn = New DataColumn("loc", Type.GetType("System.String"))
            Dim bin_no As DataColumn = New DataColumn("bin_no", Type.GetType("System.String"))
            Dim qty_on_hand As DataColumn = New DataColumn("qty_on_hand", Type.GetType("System.Int32"))

            .Columns.Add(item_no)
            .Columns.Add(bin_no)
            .Columns.Add(qty_on_hand)
            .Columns.Add(loc)
        End With

        With dtWarehouse002
            Dim item_no As DataColumn = New DataColumn("item_no", Type.GetType("System.String"))
            Dim loc As DataColumn = New DataColumn("loc", Type.GetType("System.String"))
            Dim bin_no As DataColumn = New DataColumn("bin_no", Type.GetType("System.String"))
            Dim qty_on_hand As DataColumn = New DataColumn("qty_on_hand", Type.GetType("System.Int32"))
            .Columns.Add(item_no)
            .Columns.Add(bin_no)
            .Columns.Add(qty_on_hand)
            .Columns.Add(loc)
        End With

        With dtTransactionHistory
            Dim ord_count As DataColumn = New DataColumn("ord_count", Type.GetType("System.Int32"))
            Dim doc_type As DataColumn = New DataColumn("doc_type", Type.GetType("System.String"))
            Dim from_to As DataColumn = New DataColumn("from_to", Type.GetType("System.String"))
            Dim ord_no As DataColumn = New DataColumn("ord_no", Type.GetType("System.String"))
            Dim bin_no As DataColumn = New DataColumn("bin_no", Type.GetType("System.String"))
            Dim quantity As DataColumn = New DataColumn("quantity", Type.GetType("System.Int32"))
            Dim trx_eff_dt As DataColumn = New DataColumn("trx_eff_dt", Type.GetType("System.String"))
            .Columns.Add(ord_count)
            .Columns.Add(doc_type)
            .Columns.Add(from_to)
            .Columns.Add(ord_no)
            .Columns.Add(bin_no)
            .Columns.Add(quantity)
            .Columns.Add(trx_eff_dt)
        End With

    End Sub

    Private Sub LoadDataGridViews()
        LoadItemDataGridView()
        LoadFactoryDataGridView()
        LoadWarehouseDataGridView()
        LoadTransactionLogDataGridView()
        LoadTransactionHistoryDataGridView()
        LoadWorkingArea(TransferType.Transfer)
    End Sub

    Private Sub LoadPicBoxTags()
        pbxIssue.Tag = TransferType.Issue.ToString
        pbxReceipt.Tag = TransferType.Receipt.ToString
        pbxTransfer.Tag = TransferType.Transfer.ToString
    End Sub

    Public Function LoadCurrentPeriod() As String
        'NOTE: This returns the Current Period for IM module, IM module is 8.  It is checked before each transaction
        '      prd_key = Period Key, or the Period we are in which can be 1, 2, 3, 4, 5, 6, 7, 8, 9, A, B, C (JAN = 1, FEB = 2,... OCT = A, NOV = B DEC = C)
        Dim sIMModuleNo As String = "8"
        Dim sSQL As String = "select current_prd, *  from SYPRDFIL_SQL where prd_key = '" & sIMModuleNo & "'"
        Dim CurrentPeriod As String = Convert.ToString(cBusObj.ExecuteSQLScalar(sSQL, cn)).Trim
        Dim CurrentPeriodAlias As String = ""
        Dim labelText As String = ""
        Select Case CurrentPeriod
            Case "A"
                CurrentPeriodAlias = "10"
            Case "B"
                CurrentPeriodAlias = "11"
            Case "C"
                CurrentPeriodAlias = "12"
            Case Else
                CurrentPeriodAlias = CurrentPeriod
        End Select
        labelText = "(Current Period is " & CurrentPeriodAlias & ")"
        lblCurrentPeriod.Text = labelText

        Return labelText

    End Function

    Private Sub LoadSessionData()
        With sessiondata
            .SessionID = LoadSessionID()
            If .SessionID = 0 Then .SessionID = 1
            .CreateDate = cBusObj.GetMacolaDate(Now.Year.ToString, Format(Now.Month.ToString.PadLeft(2, "0")), _
                                                Format(Now.Day.ToString.PadLeft(2, "0")))
            .TransactionID = 1
        End With
    End Sub

    Private Function LoadSessionID() As Integer
        Dim id As Object

        id = cBusObj.ExecuteSQLScalar("Select max(session_id) + 1 from IMREPLOG_MAS", cn)
        If id Is DBNull.Value Then
            id = 1
        End If

        Return id
    End Function

    Private Sub LoadData(AppLoading As Boolean, Optional itm As String = "")

        Dim sSQL As String

        uidata.ItemNo = txtM.Text.Trim & txtItemSearch.Text.Trim
        bPaintDGVBorders = True
        If bAppLoading Then
            'add the items ....
            sSQL = "Select itm.item_no, itm.item_desc_1 from IMITMIDX_SQL itm where SUBSTRING(itm.item_no, 1, 1) = 'M'"
            dtAllItems = cBusObj.ExecuteSQLDataTable(sSQL, "Items", cn)

            'add Warehouse Bins ...
            sSQL = "Select Distinct item_no, lTrim(rTrim(loc)) as loc, lTrim(rTrim(bin_no)) as bin_no, qty_on_hand  from IMINVBIN_SQL Where loc = '00" & CStr(Locations.Warehouse) & "'  and SUBSTRING(Item_no, 1, 1) = 'M'"
            dtAllWarehouseBins = cBusObj.ExecuteSQLDataTable(sSQL, "WarehouseBin", cn)
            dvWareHouse = dtAllWarehouseBins.DefaultView

            'add Factory Bins ...
            sSQL = "Select Distinct item_no, lTrim(rTrim(loc)) as loc, lTrim(rTrim(bin_no)) as bin_no, qty_on_hand  from IMINVBIN_SQL Where loc = '00" & CStr(Locations.Factory) & "' and SUBSTRING(Item_no, 1, 1) = 'M' "
            dtAllFactoryBins = cBusObj.ExecuteSQLDataTable(sSQL, "FactoryBin", cn)
            'dvFactory is for the Factory DataGridView
            dvFactory = dtAllFactoryBins.DefaultView

            'dvFactoryForCombo is for the Combo box that is in the dvWorkingArea for Factory Bins.  It must have all bins at all times.  
            'add Warehouse Bins for Combo ...
            sSQL = "Select Distinct '' as item_no, lTrim(rTrim(loc)) as loc, lTrim(rTrim(bin_no)) as bin_no, 0 as qty_on_hand from IMINVBIN_SQL Where loc = '00" & CStr(Locations.Factory) & "' "
            dtFactoryForCombo = cBusObj.ExecuteSQLDataTable(sSQL, "FactoryBinForCombo", cn)
            dvFactoryForCombo = dtFactoryForCombo.DefaultView

            bAppLoading = False
        Else
            dtItem = dtAllItems.Clone

            For Each rw As DataRow In dtAllItems.Select("Trim(item_no) = '" & itm & "'")
                dtItem.ImportRow(rw)
            Next

            'Filter the DataViews for the DataGridView Comboboxes 
            dvFactory.RowFilter = "Trim(item_no) = '" & itm & "'"
            If dvFactory.Count = 0 Then
                sSQL = "Select Distinct '" & itm & "' as item_no, '00" & CStr(Locations.Factory) & "'" & " as loc, '' as bin_no, 0 as qty_on_hand  from IMINVBIN_SQL Where loc = '00" & CStr(Locations.Factory) & "' "
                dtAllFactoryBins = cBusObj.ExecuteSQLDataTable(sSQL, "FactoryBin", cn)
                dvFactory = dtAllFactoryBins.DefaultView
            End If

            dvWareHouse = dtAllWarehouseBins.DefaultView
            dvWareHouse.RowFilter = "item_no = '" & itm & "'"
            If dvWareHouse.Count = 0 Then
                sSQL = "Select Distinct '" & itm & "' as item_no, '00" & CStr(Locations.Warehouse) & "'" & " as loc, '' as bin_no, 0 as qty_on_hand  from IMINVBIN_SQL Where loc = '00" & CStr(Locations.Warehouse) & "' "
                Dim dtWhse = cBusObj.ExecuteSQLDataTable(sSQL, "FactoryBin", cn)
                dvWareHouse = dtWhse.DefaultView
            End If

            'Populate the dtWareHouse002 and dtFactory001 for the respective DataGridViews where they display
            dtWarehouse002 = dtAllWarehouseBins.Clone

            For Each rw As DataRow In dtAllWarehouseBins.Select("item_no = '" & itm & "' and loc = '00" & CStr(Locations.Warehouse) & "' ")
                dtWarehouse002.ImportRow(rw)
            Next

            dtFactory001 = dtAllFactoryBins.Clone

            For Each rw As DataRow In dtAllFactoryBins.Select("item_no = '" & itm & "' and loc = '00" & CStr(Locations.Factory) & "' ")
                dtFactory001.ImportRow(rw)
            Next

        End If

        Dim rwdiff As Integer = 0

        If dtWarehouse002.Rows.Count + dtFactory001.Rows.Count = 0 Then
            Dim arRow() As Object = {itm.ToUpper, " - ", 0}
            dtWarehouse002.Rows.Add(arRow)
            arRow(0) = " - "
            dtFactory001.Rows.Add(arRow)
        ElseIf dtWarehouse002.Rows.Count = 0 Then
            Dim arRow() As Object = {itm.ToUpper, " - ", 0}
            dtWarehouse002.Rows.Add(arRow)
        End If

        dgvItems.DataSource = dtItem
        dgvFactory001.DataSource = dtFactory001
        dgvWarehouse002.DataSource = dtWarehouse002

        UnselectDataGridViewRows(dgvItems)
        UnselectDataGridViewRows(dgvFactory001)
        UnselectDataGridViewRows(dgvWarehouse002)

        rwdiff = dgvWarehouse002.Rows.Count - dgvFactory001.Rows.Count
        If rwdiff > 0 Then
            Dim arRow() As Object = {itm, "001", " ", 0}
            For i As Integer = 0 To rwdiff - 1
                dtFactory001.Rows.Add(arRow)
            Next
        End If
        rwdiff = dgvWarehouse002.Rows.Count - dgvItems.Rows.Count
        If rwdiff > 0 Then
            Dim arRow() As Object = {"", ""}
            For i As Integer = 0 To rwdiff - 1
                dtItem.Rows.Add(arRow)
            Next
        End If

    End Sub

    Private Sub PopulateTransactionDate() Handles fTransactionDate1.SetTransactionDate
        'PopulateTransactionDate is called by fTransactionDate on Startup
        TransactionEffectiveDate = fTransactionDate1.TransactionDateDateTimePicker.Value
        dtpTransactionDate.Value = TransactionEffectiveDate
        TransEffectiveDate = cBusObj.GetMacolaDate(TransactionEffectiveDate)
        txtItemSearch.Focus()
    End Sub

    Private Sub PopulateTransferToBin() Handles fFactoryBinSelect1.SetFactoryBin
        TransferToBin = fFactoryBinSelect.cboDefaultFactoryBin.Text
        cboDefaultFactoryBin.Text = TransferToBin
    End Sub

#End Region

#Region "  REGION - Load DataGridViews   "

    Private Sub LoadItemDataGridView()

        'DataGridView ...
        With dgvItems
            Me.Panel2.Controls.Add(dgvItems)
            Me.Panel2.Controls.Add(Me.TextBox2)
            .Dock = DockStyle.Fill
            '.Location = New System.Drawing.Point(0, 24)

            .AllowUserToResizeRows = False
            .AllowUserToAddRows = False
            .RowHeadersVisible = False
            .DataSource = Nothing
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .ColumnHeadersVisible = True
            .DefaultCellStyle.Font = New Font("Segoe UI", 9)
            .ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 10)
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
            .EnableHeadersVisualStyles = False


            If .RowCount > 0 Then .Rows.Clear()
            If .ColumnCount > 0 Then .Columns.Clear()

            Dim chkColumn As New DataGridViewCheckBoxColumn
            With chkColumn
                .DataPropertyName = "selected"
                .Name = "Selected"
                .HeaderText = sCheckMark3
                '.HeaderText = sGlyphUp
                .MinimumWidth = 22
                .ToolTipText = "Click to Sort : Dbl-Click to Check/UnCheck All"
                .Width = 22

                .SortMode = DataGridViewColumnSortMode.NotSortable
                .Visible = False
            End With
            .Columns.Add(chkColumn)

            With .Columns(.Columns.Add("ItemNo", "Item No"))
                .MinimumWidth = 100
                .ToolTipText = "Order No"
                .DataPropertyName = "item_no"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 45
                .ReadOnly = True
            End With

            With .Columns(.Columns.Add("ItemDescription", "Description"))
                .MinimumWidth = 100
                .DataPropertyName = "item_desc_1"
                .ToolTipText = "ItemDescription"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .WrapMode = DataGridViewTriState.True
                End With
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                '.Width = 50
                .ReadOnly = True
            End With

            'Turn of sorting
            Dim i As Integer
            For i = 0 To .Columns.Count - 1
                .Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            Next i
        End With
    End Sub

    Private Sub LoadFactoryDataGridView()

        'DataGridView ...
        With dgvFactory001
            .AllowUserToResizeRows = False
            .AllowUserToAddRows = False
            .RowHeadersVisible = False
            .DataSource = Nothing
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect

            .DefaultCellStyle.Font = New Font("Segoe UI", 9)
            .ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 10)
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
            .EnableHeadersVisualStyles = False

            If .RowCount > 0 Then .Rows.Clear()
            If .ColumnCount > 0 Then .Columns.Clear()

            With .Columns(.Columns.Add("ItemNo", "Item No"))
                .MinimumWidth = 100
                .ToolTipText = "Order No"
                .DataPropertyName = "item_no"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 45
                .ReadOnly = True
                .Visible = False
            End With

            With .Columns(.Columns.Add("Bin", "Bin"))
                .DataPropertyName = "bin_no"
                .ToolTipText = "Bin Number"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 80
                .MinimumWidth = 50
                .ReadOnly = True
            End With

            With .Columns(.Columns.Add("Qty", "Qty"))
                .ToolTipText = "Quantity On Hand"
                .HeaderCell.ToolTipText = .ToolTipText
                .DataPropertyName = "qty_on_hand"
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    .Format = "N0"
                End With
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 65
                .MinimumWidth = 55
                .ReadOnly = True
            End With

            With .Columns(.Columns.Add("Loc", "Loc"))
                .DataPropertyName = "loc"
                .ToolTipText = "Warehouse Location"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 80
                .MinimumWidth = 50
                .ReadOnly = True
            End With

            'Turn off sorting
            Dim i As Integer
            For i = 0 To .Columns.Count - 1
                .Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            Next i
        End With

    End Sub

    Private Sub LoadWarehouseDataGridView()

        'DataGridView ...
        With dgvWarehouse002
            .AllowUserToResizeRows = False
            .AllowUserToAddRows = False
            .RowHeadersVisible = False
            .DataSource = Nothing
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect

            .DefaultCellStyle.Font = New Font("Segoe UI", 9)
            .ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 10)
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
            .EnableHeadersVisualStyles = False

            If .RowCount > 0 Then .Rows.Clear()
            If .ColumnCount > 0 Then .Columns.Clear()

            With .Columns(.Columns.Add("ItemNo", "Item No"))
                .ToolTipText = "Item No"
                .DataPropertyName = "item_no"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .ReadOnly = True
                .Visible = False
            End With

            With .Columns(.Columns.Add("Bin", "Bin"))
                .DataPropertyName = "bin_no"
                .ToolTipText = "Bin Number"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 80
                .MinimumWidth = 50
                .ReadOnly = True
            End With

            With .Columns(.Columns.Add("Qty", "Qty"))
                .ToolTipText = "Quantity On Hand"
                .HeaderCell.ToolTipText = .ToolTipText
                .DataPropertyName = "qty_on_hand"
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    .Format = "N0"
                End With
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 65
                .MinimumWidth = 55
                .ReadOnly = True
            End With

            With .Columns(.Columns.Add("Loc", "Loc"))
                .DataPropertyName = "loc"
                .ToolTipText = "Warehouse Location"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 80
                .MinimumWidth = 50
                .ReadOnly = True
            End With

            ''Turn off sorting
            'Dim i As Integer
            'For i = 0 To .Columns.Count - 1
            '    .Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            'Next i
        End With

    End Sub

    Private Sub LoadTransactionLogDataGridView()

        'Build the DataGridView ...
        With dgvTransactionLog
            .AllowUserToResizeRows = False
            .AllowUserToAddRows = False
            .RowHeadersVisible = False
            .DataSource = Nothing
            .DefaultCellStyle.Font = New Font("Segoe UI", 9)
            .ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 10)
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
            .EnableHeadersVisualStyles = False

            If .RowCount > 0 Then .Rows.Clear()
            If .ColumnCount > 0 Then .Columns.Clear()

            With .Columns(.Columns.Add("TranType", "Type"))
                .MinimumWidth = 40
                .ToolTipText = "Transfer Type"
                .DataPropertyName = "tran_type"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                .Width = 65
                .ReadOnly = True
            End With

            With .Columns(.Columns.Add("ItemNo", "Item"))
                .MinimumWidth = 45
                .ToolTipText = "Item No"
                .DataPropertyName = "item_no"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                .Width = 50
                .ReadOnly = True
            End With

            With .Columns(.Columns.Add("LocFrom", "From"))
                .MinimumWidth = 45
                .ToolTipText = "From Location"
                .DataPropertyName = "loc_from"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                .Width = 40
                .ReadOnly = True
            End With

            With .Columns(.Columns.Add("BinFrom", "Bin"))
                .MinimumWidth = 45
                .ToolTipText = "From Bin"
                .DataPropertyName = "bin_from"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                .Width = 40
                .ReadOnly = True
            End With

            'Qty is the number of items at Bin_FROM and if populated from the Drag Drop from Factory or Warehouse Datagridviews
            ' the column is NOT visible, and thus maintains the original dragged qty and can be used along with TranactionQty below 
            ' for determining the qty remaining after the transaction ...
            With .Columns(.Columns.Add("BinFromQtyOnHand", "Prev Qty"))
                .ToolTipText = "Bin From Quantity On Hand"
                .HeaderCell.ToolTipText = .ToolTipText
                .DataPropertyName = "bin_from_qty_on_hand"
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    .Format = "N0"
                End With
                .AutoSizeMode = .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                .Width = 45
                .ReadOnly = True
                .Visible = True
            End With

            'Transaction Qty is Qty is the suggested qty to move into each Bin_To and is based on the last Transaction
            ' for the combination of Item, Location when it was transferred out (as a Bin_From).  Normally the qty should
            ' be the same, but it can be changed in the grid. 
            With .Columns(.Columns.Add("BinFromTransQty", "Trx Qty"))
                .ToolTipText = "Bin From Transaction Quantity"
                .HeaderCell.ToolTipText = .ToolTipText
                .DataPropertyName = "bin_from_trx_qty"
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    .Format = "N0"
                End With
                .AutoSizeMode = .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                '.Width = 55
                .ReadOnly = True
                .Visible = False
            End With

            With .Columns(.Columns.Add("LocTo", "To"))
                .MinimumWidth = 40
                .ToolTipText = "To Location"
                .DataPropertyName = "loc_to"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                .Width = 40
                .ReadOnly = True
            End With

            With .Columns(.Columns.Add("BinTo", "Bin"))
                .MinimumWidth = 40
                .ToolTipText = "To Bin"
                .DataPropertyName = "bin_to"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                .Width = 40
                .ReadOnly = True
            End With

            With .Columns(.Columns.Add("BinToQtyOnHand", "Prev Qty"))
                .ToolTipText = "Bin To Quantity On Hand"
                .HeaderCell.ToolTipText = .ToolTipText
                .DataPropertyName = "bin_to_qty_on_hand"
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    .Format = "N0"
                End With
                '.AutoSizeMode = .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                .Width = 45
                .ReadOnly = True
                .Visible = True
            End With

            'Transaction Qty is Qty repeated, for use, see Qty above ^, up there ^... 
            ' meant to be a qty starting point and visible so the user can edit it
            With .Columns(.Columns.Add("BinToTransQty", "Trx Qty"))
                .ToolTipText = "Bin To Transaction Quantity"
                .HeaderCell.ToolTipText = .ToolTipText
                .DataPropertyName = "bin_to_trx_qty"
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    .Format = "N0"
                End With
                '.AutoSizeMode = .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                '.Width = 55
                .ReadOnly = True
                .Visible = False
            End With

            'Transaction Qty is Qty repeated, for use, see Qty above ^, up there ^... 
            ' meant to be a qty starting point and visible so the user can edit it
            With .Columns(.Columns.Add("QtyToMove", "Trx Qty"))
                .ToolTipText = "Quantity To Move"
                .HeaderCell.ToolTipText = .ToolTipText
                .DataPropertyName = "qty_to_move"
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    .Format = "N0"
                End With
                '.AutoSizeMode = .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                .Width = 55
                .ReadOnly = True
            End With

            With .Columns(.Columns.Add("SessionID", "Ses ID"))
                .ToolTipText = "Session ID"
                .HeaderCell.ToolTipText = .ToolTipText
                .DataPropertyName = "session_id"
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    .Format = "N0"
                End With
                '.AutoSizeMode = .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                .Width = 40
                .ReadOnly = True
                .Visible = True
            End With

            With .Columns(.Columns.Add("TransactionID", "Trn ID"))
                .ToolTipText = "Transaction ID"
                .HeaderCell.ToolTipText = .ToolTipText
                .DataPropertyName = "tran_id"
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    .Format = "N0"
                End With
                '.AutoSizeMode = .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                .Width = 40
                .ReadOnly = True
                .Visible = True
            End With

            With .Columns(.Columns.Add("CreateDate", "Create Date"))
                .ToolTipText = "Create Date"
                .HeaderCell.ToolTipText = .ToolTipText
                .DataPropertyName = "create_dt"
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    '.Format = "N0"
                End With
                '.AutoSizeMode = .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                .Width = 60
                .ReadOnly = True
                .Visible = True
            End With

            'Turn off sorting
            'Dim i As Integer
            'For i = 0 To .Columns.Count - 1
            '    .Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            'Next i
        End With

    End Sub

    Private Sub LoadTransactionHistoryDataGridView()
        With dgvTransactionHistory

            .AllowUserToResizeRows = False
            .AllowUserToAddRows = False
            .RowHeadersVisible = False
            .RowHeadersWidth = 22
            .DataSource = Nothing
            .DefaultCellStyle.Font = New Font("Segoe UI", 9)
            .ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 10)
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
            .EnableHeadersVisualStyles = False
            .EditMode = DataGridViewEditMode.EditOnEnter

            If .RowCount > 0 Then .Rows.Clear()
            If .ColumnCount > 0 Then .Columns.Clear()

            With .Columns(.Columns.Add("OrderCount", ""))
                .ToolTipText = ""
                .HeaderCell.ToolTipText = .ToolTipText
                .DataPropertyName = "ord_count"
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    .Format = "N0"
                End With
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                .Width = 30
                .ReadOnly = False
                .Visible = False
            End With

            With .Columns(.Columns.Add("Type", ""))
                .MinimumWidth = 25
                .ToolTipText = "Transaction Type"
                .DataPropertyName = "doc_type"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    .Font = New Font("Segoe UI", 9)
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 25
                .ReadOnly = True
                .Visible = False
            End With

            With .Columns(.Columns.Add("TransTypeDescription", "Desc"))
                .MinimumWidth = 75
                .ToolTipText = "Transaction Type Description"
                .DataPropertyName = "from_to"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .WrapMode = DataGridViewTriState.True
                    .Font = New Font("Segoe UI", 9)
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 75
                .ReadOnly = True
                .DisplayIndex = 1
            End With

            With .Columns(.Columns.Add("Order No", "Order No"))
                .MinimumWidth = 65
                .ToolTipText = "Order Number"
                .DataPropertyName = "ord_no"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .WrapMode = DataGridViewTriState.True
                    .Font = New Font("Segoe UI", 9)
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 65
                .ReadOnly = True
                .DisplayIndex = 0
            End With

            With .Columns(.Columns.Add("Bin No", "Bin"))
                .MinimumWidth = 40
                .ToolTipText = "Bin Number"
                .DataPropertyName = "bin_no"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleRight
                    .WrapMode = DataGridViewTriState.True
                    .Font = New Font("Segoe UI", 9)
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 40
                .ReadOnly = True
            End With

            With .Columns(.Columns.Add("Quantity", "Qty"))
                .ToolTipText = "Transaction Quantity"
                .HeaderCell.ToolTipText = .ToolTipText
                .DataPropertyName = "quantity"
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    .Format = "N0"
                End With
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                .Width = 40
                .ReadOnly = False
                .Visible = True
            End With

            With .Columns(.Columns.Add("TransDate", "Trx Date"))
                .MinimumWidth = 70
                .ToolTipText = "Transaction Date"
                .DataPropertyName = "trx_dt"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .WrapMode = DataGridViewTriState.True
                    .Font = New Font("Segoe UI", 9)
                End With
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 70
                .ReadOnly = True
            End With

        End With

    End Sub

    Private Sub LoadWorkingArea(TransferType As Integer)

        'Build the DataGridView ...
        With dgvWorkingArea
            .AllowUserToResizeRows = False
            .AllowUserToAddRows = False
            .RowHeadersVisible = True
            .RowHeadersWidth = 22
            .DataSource = Nothing
            .DefaultCellStyle.Font = New Font("Segoe UI", 9)
            .ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 10)
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
            .EnableHeadersVisualStyles = False
            .EditMode = DataGridViewEditMode.EditOnEnter

            If .RowCount > 0 Then .Rows.Clear()
            If .ColumnCount > 0 Then .Columns.Clear()

            With .Columns(.Columns.Add("TranType", "Tran Type"))
                .MinimumWidth = 100
                .ToolTipText = "Transfer Type"
                .DataPropertyName = "tran_type"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 45
                .ReadOnly = True
            End With

            With .Columns(.Columns.Add("ItemNo", "Item No"))

                .MinimumWidth = 100
                .ToolTipText = "Item No"
                .DataPropertyName = "item_no"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleLeft
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 45
                .ReadOnly = False
            End With

            With dgvComboLocFrom
                .MinimumWidth = 45
                .Name = "TargetLocationFrom"
                .HeaderText = "From"
                .ToolTipText = "From Loc"
                .DataPropertyName = "target_loc_from"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                With dgvComboLocFrom
                    .Items.Add("001")
                    .Items.Add("002")
                End With
                .Width = 70
                .ReadOnly = False
            End With
            .Columns.Add(dgvComboLocFrom)

            With dgvComboBinFirst
                .Name = "BinFirst"
                .HeaderText = "Bin From"
                .MinimumWidth = 35
                .DataPropertyName = "bin_no"
                .ToolTipText = "Bin Warehouse"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 80
                .ReadOnly = False

            End With
            .Columns.Add(dgvComboBinFirst)

            With .Columns(.Columns.Add("Direction", ""))
                .MinimumWidth = 30
                .ToolTipText = "Direction"
                .DataPropertyName = "direction"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    .Font = New Font("Segoe UI", 7)
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 35
                .ReadOnly = True
            End With

            With dgvComboLocTo
                .Name = "TargetLocationTo"
                .MinimumWidth = 45
                .HeaderText = "To"
                .ToolTipText = "To Location"
                .DataPropertyName = "target_loc_to"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                With dgvComboLocTo
                    .Items.Add("001")
                    .Items.Add("002")
                End With
                .Width = 70
                .ReadOnly = False
            End With
            .Columns.Add(dgvComboLocTo)

            With dgvComboBinSecond
                .Name = "BinSecond"
                .HeaderText = "Bin To"
                .MinimumWidth = 35
                .DataPropertyName = "bin_no"
                .ToolTipText = "Bin Factory"
                .HeaderCell.ToolTipText = .ToolTipText
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 80
                .ReadOnly = False
                .AutoComplete = True
            End With
            .Columns.Add(dgvComboBinSecond)

            'Qty is the number of items at Bin_FROM and if populated from the Drag Drop from Factory or Warehouse Datagridviews
            ' the column is NOT visible, and thus maintains the original dragged qty and can be used along with TranactionQty below 
            ' for determining the qty remaining after the transaction ...
            With .Columns(.Columns.Add("BinFromQtyOnHand", "Bin From On Hand"))
                .ToolTipText = "Bin From Quantity On Hand"
                .HeaderCell.ToolTipText = .ToolTipText
                .DataPropertyName = "bin_from_qty_on_hand"
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    .Format = "N0"
                    .BackColor = SystemColors.ButtonFace
                End With
                '.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Width = 105
                .MinimumWidth = 75
                .ReadOnly = True
                'TCC made change here to True, change back to False
                .Visible = True
            End With

            'Transaction Qty is Qty is the suggested qty to move into each Bin_To and is based on the last Transaction
            ' for the combination of Item, Location when it was transferred out (as a Bin_From).  Normally the qty should
            ' be the same, but it can be changed in the grid. 
            With .Columns(.Columns.Add("BinFromTransQty", "BinFromTransQty"))
                .ToolTipText = "Bin From Transaction Quantity"
                .HeaderCell.ToolTipText = .ToolTipText
                .DataPropertyName = "bin_from_transqty"
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    .Format = "N0"
                End With
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                '.Width = 55
                .ReadOnly = False
                'TCC made change here to True, change back to False
                .Visible = False
            End With

            With .Columns(.Columns.Add("BinToQtyOnHand", "BinToQtyOnHand"))
                .ToolTipText = "Bin To Quantity On Hand"
                .HeaderCell.ToolTipText = .ToolTipText
                .DataPropertyName = "bin_to_qty_on_hand"
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    .Format = "N0"
                End With
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                '.Width = 55
                .ReadOnly = False
                'TCC made change here to True, change back to False
                .Visible = False
            End With

            'Transaction Qty is Qty repeated, for use, see Qty above ^, up there ^... 
            ' meant to be a qty starting point and visible so the user can edit it
            With .Columns(.Columns.Add("BinToTransQty", "BinToTransQty"))
                .ToolTipText = "Bin To Transaction Quantity"
                .HeaderCell.ToolTipText = .ToolTipText
                .DataPropertyName = "bin_to_transqty"
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    .Format = "N0"
                End With
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                '.Width = 55
                .ReadOnly = False
                'TCC made change here to True, change back to False
                .Visible = False
            End With

            'Transaction Qty is Qty repeated, for use, see Qty above ^, up there ^... 
            ' meant to be a qty starting point and visible so the user can edit it
            With .Columns(.Columns.Add("QtyToMove", "Quantity To Move"))
                .ToolTipText = "Quantity To Move"
                .HeaderCell.ToolTipText = .ToolTipText
                .DataPropertyName = "qty_to_move"
                With .DefaultCellStyle
                    .Alignment = DataGridViewContentAlignment.MiddleCenter
                    .WrapMode = DataGridViewTriState.True
                    .Format = "N0"
                End With
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .MinimumWidth = 75
                '.Width = 55
                .ReadOnly = False
            End With

        End With

    End Sub

#End Region

#Region "  REGION - Load DataTables  "

    Private Function LoadWorkingAreaDataTable() As DataTable
        Dim dt As New DataTable

        ''create the datatable 
        dt.Columns.Add("trantype", GetType(String))
        dt.Columns.Add("itemno", GetType(String))
        dt.Columns.Add("binfrom", GetType(String))
        dt.Columns.Add("targetloc_from", GetType(String))
        dt.Columns.Add("bin_from_qty_on_hand", GetType(Integer))
        dt.Columns.Add("bin_from_transqty", GetType(Integer))
        dt.Columns.Add("binto", GetType(String))
        dt.Columns.Add("targetlocto", GetType(String))
        dt.Columns.Add("bin_to_qty_on_hand", GetType(Integer))
        dt.Columns.Add("bin_to_transqty", GetType(Integer))

        Return dt
    End Function

#End Region

#Region "  REGION - Button, Textbox & Menu Control Events   "

#Region "  REGION - Context Menu & Items   "



    Private Sub dgvWorkingArea_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgvWorkingArea.MouseDown
        Dim dgv As DataGridView = CType(sender, DataGridView)
        If e.Button = Windows.Forms.MouseButtons.Right Then
            'Dim ht As DataGridView.HitTestInfo
            ht = dgv.HitTest(e.X, e.Y)
            hitContextMenu = dgv.HitTest(e.X, e.Y)

            'to set current row programatically, Because current row can be multiple rows
            'when row selection is set to MultiSelect, it cannot be set where the little
            'black arrow will move to that row.  But setting the Cell to CurrentCell, there is only
            'one currentcell, not multiples.  So, set Cell first, then row, and the little black 
            'arrow will move. 
            Try
                With dgv
                    .ClearSelection()
                    .CurrentCell = .Rows(ht.RowIndex).Cells(ht.ColumnIndex)
                    .Rows(ht.RowIndex).Selected = True
                End With
            Catch ex As Exception

            End Try

        End If

    End Sub

    Private Sub mnuRemoveRow_Click(sender As System.Object, e As System.EventArgs) Handles mnuRemoveRow.Click
        If dgvFromEditControlShowing IsNot Nothing Then
            dgvFromEditControlShowing.Rows.Remove(dgvFromEditControlShowing.CurrentRow)
        Else
            Dim cms As ContextMenuStrip = CType(CType(sender, ToolStripMenuItem).GetCurrentParent, ContextMenuStrip)
            Dim dgv As DataGridView = CType(cms.SourceControl, DataGridView)
            dgv.Rows.Remove(dgv.CurrentRow)
        End If

        'Timer2 re-does the QTY Calculation on transfers after removing a row ...
        With Timer2
            .Interval = 500
            .Enabled = True
        End With

    End Sub

    Private Sub mnuCopyDown_Click(sender As System.Object, e As System.EventArgs) Handles mnuCopyDown.Click
        Dim dgv As DataGridView = CType(dgvWorkingArea, DataGridView)

        With dgv
            Dim rw As Integer = hitContextMenu.RowIndex
            Dim cl As Integer = hitContextMenu.ColumnIndex
            Dim itm As String = .Rows(rw).Cells("ItemNo").Value
            Dim val As String = .Rows(rw).Cells(cl).Value
            For Each row As DataGridViewRow In .Rows
                If .Rows(row.Index).Cells("ItemNo").Value = itm Then
                    .Rows(row.Index).Cells(cl).Value = val
                    Dim eventargs As New DataGridViewCellEventArgs(cl, row.Index)
                    dgvWorkingArea_CellEndEdit(dgv, eventargs)
                End If
            Next
        End With

    End Sub

    Private Sub AddRowToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles mnuAddRow.Click
        Dim dgv As DataGridView = CType(dgvWorkingArea, DataGridView)
        Dim rw As Integer = dgv.CurrentRow.Index
        Dim myQty As Integer = 0
        Dim target As Integer = 0

        DataGridViewAddRow(dgv, rw)

        With Timer6
            .Interval = 500
            .Enabled = True
        End With

    End Sub

    Private Sub DataGridViewAddRow(dgv As DataGridView, rw As Integer)
        Dim newrw As DataGridViewRow = dgv.Rows(rw).Clone

        For Each cel As DataGridViewCell In dgv.Rows(rw).Cells
            newrw.Cells(cel.ColumnIndex).Value = cel.Value
        Next

        dgv.Rows.AddRange(newrw)

    End Sub

#End Region

#Region "  REGION - Button Events  "

    Private Sub btnAddMold_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddMold.Click
        Dim itm As String = txtM.Text.Trim & txtItemSearch.Text.Trim
        If itm = "M" Then Exit Sub

        ValidateTransactionDate() ' this check is only to ensure the IM System Period is always correct before a transaction can begin

        bAppLoading = True
        If ValidateData() = True Then
            Clear()
            LoadData(bAppLoading, itm)
            LoadData(bAppLoading, itm)
        End If

        LoadTrasactionHistory(itm)

    End Sub

    Private Sub btnAddMold_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddMold.MouseEnter
        btnAddMold.BackColor = SystemColors.ScrollBar
    End Sub

    Private Sub btnAddMold_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddMold.MouseLeave
        btnAddMold.BackColor = SystemColors.ButtonFace
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Clear()
    End Sub

    Private Sub btnClear_MouseEnter(sender As Object, e As System.EventArgs)
        Dim btn As Button = CType(sender, Button)
        Select Case btn.Name
            Case btnClearLog.Name
                btn.BackColor = Color.OrangeRed
                btn.ForeColor = Color.White
                'Case btnSearchLog.Name
                '    btn.BackColor = SystemColors.GradientInactiveCaption
                '    'btn.ForeColor = Color.White
        End Select


    End Sub

    Private Sub btnClear_MouseLeave(sender As Object, e As System.EventArgs)
        Dim btn As Button = CType(sender, Button)
        Select Case btn.Name
            Case btnClearLog.Name
                btn.BackColor = SystemColors.Control
                btn.ForeColor = SystemColors.WindowText
                'Case btnSearchLog.Name
                '    btn.BackColor = SystemColors.Control
        End Select

    End Sub

    Private Sub btnSearchLog_Click(sender As System.Object, e As System.EventArgs) Handles btnSearchLog.Click
        Dim createdate As Date = dtpRepetitiveCreateDate.Value
        Dim searchdate As Integer = cBusObj.GetMacolaDate(createdate.Year.ToString, Format(createdate.Month.ToString.PadLeft(2, "0")), _
                                                Format(createdate.Day.ToString.PadLeft(2, "0")))
        FillTransactionLog(searchdate, RepetitiveLogType.CreateDate)
    End Sub

    Private Sub btnClearLog_Click(sender As System.Object, e As System.EventArgs)
        LoadTransactionLogDataGridView()
    End Sub

    Private Sub btnQtyRefresh_Click(sender As System.Object, e As System.EventArgs)
        DoQtyCalculation()
    End Sub

    Private Sub btnMoreLess_Click(sender As System.Object, e As System.EventArgs) Handles btnMoreLess.Click
        With SplitContainer4
            If .Panel2Collapsed = True Then
                .Panel2Collapsed = False
                .SplitterDistance = (.Width / 3) * 2
                dgvTransactionHistory.Visible = True
                dgvTransactionLog.Visible = True
                lblMoreLess.Text = "...less"
                btnMoreLess.Text = sArrowTriangleRight
            Else
                .Panel2Collapsed = True
                dgvTransactionHistory.Visible = False
                dgvTransactionLog.Visible = False
                lblMoreLess.Text = "more..."
                btnMoreLess.Text = sArrowTriangleLeft
            End If
        End With
    End Sub

    Private Sub btnOKError_Click(sender As System.Object, e As System.EventArgs) Handles btnOKError.Click
        Panel16.Visible = False
    End Sub

    Private Sub btnTransferRefresh_Click(sender As System.Object, e As System.EventArgs) Handles btnTransferRefresh.Click
        CalculateQuantity()
    End Sub

    Private Sub CalculateQuantity()
        Dim txt As TextBox = CType(txtTargetQty, TextBox)
        Dim myqty As Integer = 0
        Dim target As Integer = 0
        If IsNumeric(txt.Text) Then target = Convert.ToInt32(txt.Text)

        myqty = 0

        Try
            Dim dgv As DataGridView = CType(dgvWorkingArea, DataGridView)
            myqty = CalculateDataGridViewQty(dgv)
        Catch ex As Exception

        End Try
        
        txtMyQty.Text = myqty
        txtRemainderQty.Text = myqty - target

        If txtRemainderQty.Text <> "" OrElse txtRemainderQty.Text <> "0" Then
            SetCalculatorLabel(myqty, target)
        End If
    End Sub


#End Region

#Region "  REGION - TextBox Events  "

    Private Sub txtItemSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtItemSearch.KeyDown
        Dim itm As String = txtM.Text.Trim & txtItemSearch.Text.Trim
        If itm = "M" Then Exit Sub

        If e.KeyCode = Keys.Enter Then

            ValidateTransactionDate() ' this check is only to ensure the IM System Period is always correct before a transaction can begin

            bAppLoading = True
            If ValidateData() = True Then
                Clear()
                dgvWorkingArea.Rows.Clear()
                LoadData(bAppLoading, itm) 'Why two you ask?  First is bAppLoading = True, second is bAppLoading = False
                LoadData(bAppLoading, itm)
            End If
            LoadTrasactionHistory(itm)
        End If
    End Sub

    Private Sub txtWarehouseBinsLabel_Click(sender As Object, e As System.EventArgs) Handles txtWarehouseBinsLabel.Click, _
                                            txtFactoryBinsLabel.Click
        'This Event causes all selected rows from the previous DGV to be unselected.  
        'User can hit the Label "Warehouse" or "Factory", or the DGV. 
        Dim txt As TextBox = CType(sender, TextBox)
        If txt.Name = txtFactoryBinsLabel.Name Then
            UnselectDataGridViewRows(dgvWarehouse002)
        Else
            UnselectDataGridViewRows(dgvFactory001)
        End If

    End Sub

#End Region

#End Region

#Region "  REGION - ALL DataGridView Events  "

#Region "  REGION - dgvItems Events  "

    Private Sub dgvItems_RowPrePaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPrePaintEventArgs) Handles dgvItems.RowPrePaint
        'hopefully this will allow moving the focus back to the txtItemSearch textbox after pressing btnDone        
        If bSkipPaint = True Then
            bSkipPaint = False
            e.Handled = True
        End If
        ' Do not automatically paint the focus rectangle.
        e.PaintParts = e.PaintParts And Not DataGridViewPaintParts.Focus

        ' Determine whether the cell should be painted with the  
        ' custom selection background. 
        Try
            If e.RowIndex > 0 And dgvItems.Rows(e.RowIndex).Cells(1).Value.ToString.Trim <> "" Then
                ' Calculate the bounds of the row. 
                Dim rowBounds As New Rectangle( _
                   0, e.RowBounds.Top, Me.dgvItems.Columns.GetColumnsWidth(DataGridViewElementStates.Visible), e.RowBounds.Height)
                'Rectangle (
                '           X - dgv.RowHeadersWidth if the Row Header is visible and you want to leave it out of the rectangle ...
                '           Y - Upper Left hand corner of the rectangle ...
                ' Paint the custom selection background. 
                Dim backbrush As New SolidBrush(dgvItems.ColumnHeadersDefaultCellStyle.BackColor)

                Try
                    e.Graphics.FillRectangle(backbrush, rowBounds)
                    e.PaintCells(rowBounds, DataGridViewPaintParts.All And Not DataGridViewPaintParts.ContentBackground)
                    ControlPaint.DrawBorder(e.Graphics, rowBounds, _
                                        Color.Red, 3, ButtonBorderStyle.None, _
                                        Color.Black, 2, ButtonBorderStyle.Solid, _
                                        Color.Blue, 4, ButtonBorderStyle.None, _
                                        Color.Red, 2, ButtonBorderStyle.None)
                Finally
                    backbrush.Dispose()
                End Try

                e.Handled = True
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub dgvItems_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvItems.KeyDown
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim itm As String
        Dim rws As String()
        If e.KeyCode = Keys.Delete Then
            Try
                itm = dgv.Rows(dgv.CurrentRow.Index).Cells(1).Value.ToString.Trim
                rws = RowsToRemove(dgvWarehouse002, itm)
                If Not (rws Is Nothing) Then RemoveRows(rws)

            Catch ex As Exception

            End Try

        End If
        e.Handled = True
    End Sub

    Private Sub dgvItems_CellMouseEnter(ByVal sender As Object, ByVal location As DataGridViewCellEventArgs) Handles dgvItems.CellMouseEnter
        'Gives us where the mouse is when hovering over a cell. 
        mouseLocation = location

    End Sub

    Private Sub dgvItems_CellMouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvItems.CellMouseDown
        Dim dgv As DataGridView = CType(sender, DataGridView)

        UnselectDataGridViewRows(dgv)

        Try
            With dgv
                .Rows(mouseLocation.RowIndex).Selected = True
            End With
        Catch ex As Exception

        End Try

    End Sub

#End Region

#Region "  REGION - dgvFactory001 and dgvWarehouse002 Events  "

    Private Sub dgvFactory001_RowPrePaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPrePaintEventArgs) Handles dgvFactory001.RowPrePaint, dgvWarehouse002.RowPrePaint
        ' Do not automatically paint the focus rectangle.
        e.PaintParts = e.PaintParts And Not DataGridViewPaintParts.Focus

        Dim dgv As DataGridView = CType(sender, DataGridView)

        Try
            ' Determine whether the cell should be painted with the  
            ' custom selection background. 
            If e.RowIndex > 0 And dgvItems.Rows(e.RowIndex).Cells(1).Value.ToString.Trim <> "" Then

                ' Calculate the bounds of the row. 
                Dim rowBounds As New Rectangle( _
                    0, e.RowBounds.Top, dgv.Columns.GetColumnsWidth(DataGridViewElementStates.Displayed), e.RowBounds.Height)

                ' Paint the custom selection background. 
                Dim backbrush As New SolidBrush(dgv.ColumnHeadersDefaultCellStyle.BackColor)

                Try
                    e.Graphics.FillRectangle(backbrush, rowBounds)
                    e.PaintCells(rowBounds, DataGridViewPaintParts.All And Not DataGridViewPaintParts.ContentBackground)
                    ControlPaint.DrawBorder(e.Graphics, rowBounds, _
                                        Color.Red, 3, ButtonBorderStyle.None, _
                                        Color.Black, 2, ButtonBorderStyle.Solid, _
                                        Color.Blue, 4, ButtonBorderStyle.None, _
                                        Color.Red, 2, ButtonBorderStyle.None)
                Finally
                    backbrush.Dispose()
                End Try
                e.Handled = True
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub dgvFactory001_ClientSizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvFactory001.ClientSizeChanged
        Dim dgv1 As DataGridView = CType(sender, DataGridView)
        Dim dgv As DataGridView = dgvTransactionLog
        Dim pnl As Panel = Panel5

        Dim wdth As Integer
        wdth = dgvWarehouse002.Width + dgvItems.Width

        pnl.Left = wdth + 6
        pnl.Width = dgv1.Width - 6
    End Sub

    Private Sub dgvWarehouse002_Click(sender As Object, e As System.EventArgs) Handles dgvWarehouse002.Click,
                                         dgvFactory001.Click
        Dim dgv As DataGridView = CType(sender, DataGridView)
        If dgv.Name = dgvWarehouse002.Name Then
            UnselectDataGridViewRows(dgvFactory001)
        Else
            UnselectDataGridViewRows(dgvWarehouse002)
        End If

    End Sub

    Private Sub dgvFactory_RowEnter(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvFactory001.RowEnter, _
                                    dgvWarehouse002.RowEnter
        Dim dgv As DataGridView = CType(sender, DataGridView)
        If dgv.Name = dgvWarehouse002.Name Then
            UnselectDataGridViewRows(dgvFactory001)
        Else
            UnselectDataGridViewRows(dgvWarehouse002)
        End If
    End Sub

#End Region

#Region "  REGION - dgvWorkingArea Events  "

    Private Sub dgvWorkingArea_CellEnter(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWorkingArea.CellEnter
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim dgvtranstype As String = dgv.Rows(e.RowIndex).Cells("TranType").Value.ToString.Trim
        Dim errtxt As String = ""
        Dim bin As String = ""

        With dgv
            Select Case dgvtranstype
                Case TransferType.Transfer.ToString
                    If Not ValidataDataGridViewQty(dgv, e.RowIndex) Then
                        bin = .Rows(e.RowIndex).Cells("BinFirst").Value.ToString.Trim
                        errtxt = "The quantity is more than available in Bin " & bin
                        lblQtyError.Text = errtxt
                        Panel16.Visible = True
                        lblQtyCalculatorAlert.Visible = True
                        Exit Sub
                    Else
                        errtxt = ""
                        bin = ""
                        Panel16.Visible = False
                    End If
                    btnTransferRefresh.Enabled = True
                    txtTargetQty.Enabled = True
                Case TransferType.Receipt.ToString
                    Panel16.Visible = False
                    txtTargetQty.Enabled = False
                    txtMyQty.Text = ""
                    txtRemainderQty.Text = ""
                    lblQtyCalculatorAlert.Visible = False
                    btnTransferRefresh.Enabled = False
                Case TransferType.Issue.ToString
                    Panel16.Visible = False
                    txtTargetQty.Enabled = False
                    txtMyQty.Text = ""
                    txtRemainderQty.Text = ""
                    lblQtyCalculatorAlert.Visible = False
                    btnTransferRefresh.Enabled = False
            End Select
        End With

    End Sub

    Private Sub dgvWorkingArea_CellMouseDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvWorkingArea.CellMouseDoubleClick
        Dim dgv As DataGridView = CType(sender, DataGridView)
        dgv.EndEdit()
    End Sub

    Private Sub dgvWorkingArea_DataError(sender As Object, e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvWorkingArea.DataError
        e.Cancel = True
    End Sub

#End Region

#End Region

#Region "  REGION - Timers  "

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False

        'Setup the controls ...
        txtTargetQty.Focus()

        UnselectDataGridViewRows(dgvItems)
        UnselectDataGridViewRows(dgvFactory001)
        UnselectDataGridViewRows(dgvWarehouse002)

        With SplitContainer4
            .Panel2Collapsed = True
        End With

        'setup the alternating row color for future use (used with TransactionHistory DataGridView)
        styleAlternatingRowColor = New DataGridViewCellStyle
        styleAlternatingRowColor.BackColor = Color.LightGoldenrodYellow

        btnAddMold.Text = sArrowTriangleRight
        btnMoreLess.Text = sArrowTriangleLeft
        btnClearLog.Text = sXLargeCrossMark

        dgvTransactionHistory.Visible = False
        dgvTransactionLog.Visible = False
        lblQtyCalculatorAlert.Visible = False
        btnTransferRefresh.Enabled = False

        For i As Integer = 0 To dvFactoryForCombo.Count - 1
            cboDefaultFactoryBin.Items.Add(dvFactoryForCombo.Item(i)(2).ToString.Trim)
            fFactoryBinSelect.cboDefaultFactoryBin.Items.Add(dvFactoryForCombo.Item(i)(2).ToString.Trim)
        Next

        ''This code centers items in the ComboBox
        'Dim highestlength As Integer = 0
        'For i As Integer = 0 To .Items.Count - 1
        '    If .Items(i).ToString.Length > highestlength Then
        '        highestlength = .Items(i).ToString.Length
        '    End If
        'Next

        'For i As Integer = 0 To .Items.Count - 1
        '    Dim itemslength As Integer = .Items(i).ToString.Length
        '    If itemslength <> highestlength Then
        '        Dim addlength As Integer = CInt((highestlength - itemslength) / 2)
        '        Dim addstring As String = ""
        '        If CStr(highestlength / 2).Contains(".") Then
        '            addlength += 1
        '        End If
        '        For x As Integer = 1 To CInt(addlength)  'CInt(CInt(addlength) - (itemslength / 2))
        '            addstring &= " "
        '        Next
        '        .Items(i) = addstring & .Items(i).ToString
        '    End If

        'Next

        'finally, load the Transaction Date Form to get started
        fTransactionDate1 = New fTransactionDate
        fTransactionDate1.Show()
    End Sub

    Private Sub Timer2_Tick(sender As System.Object, e As System.EventArgs) Handles Timer2.Tick
        Timer2.Enabled = False
        Dim dgv As DataGridView = CType(dgvWorkingArea, DataGridView)
        Dim myqty As Integer = 0

        Try

            Dim target As Integer = IIf(IsNumeric(txtTargetQty.Text), Convert.ToInt32(txtTargetQty.Text), 0)
            For Each rw As DataGridViewRow In dgv.Rows
                myqty += IIf(IsNumeric(rw.Cells("QtyToMove").Value), rw.Cells("QtyToMove").Value, 0)
            Next

            txtMyQty.Text = myqty
            txtRemainderQty.Text = myqty - target
            SetCalculatorLabel(myqty, target)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Timer3_Tick(sender As System.Object, e As System.EventArgs) Handles Timer3.Tick
        Timer3.Enabled = False
        Dim errtxt As String = ""
        Dim bin As String = ""
        Dim dgv As DataGridView = CType(dgvWorkingArea, DataGridView)
        Dim rwidx As Integer = 0
        Dim dgvrowTrantype As String = ""

        With dgv
            .EndEdit()
            rwidx = .CurrentCell.RowIndex
            dgvrowTrantype = .Rows(rwidx).Cells("TranType").Value.ToString.Trim

            If dgvrowTrantype = TransferType.Transfer.ToString Then
                If Not ValidataDataGridViewQty(dgv, rwidx) Then
                    bin = .Rows(rwidx).Cells("BinFirst").Value.ToString.Trim
                    errtxt = "The quantity is more than available in Bin " & bin
                    'MsgBox(errtxt, MsgBoxStyle.OkOnly, "Quantity to Move Invalid")
                    lblQtyError.Text = errtxt
                    Panel16.Visible = True
                    Exit Sub
                End If
            Else
                errtxt = ""
                bin = ""
                Panel16.Visible = False
            End If
        End With

    End Sub

#End Region

#Region "  REGION - Methods   "

    Private Sub LoadTrasactionHistory(itm As String)
        Dim dgv As DataGridView = CType(dgvTransactionHistory, DataGridView)
        dtTransactionHistory = cBusObj.GetTransactionHistory(itm, cn)

        With dgv
            .DataSource = dtTransactionHistory
            For Each rw As DataGridViewRow In dgv.Rows
                If rw.Cells("OrderCount").Value Mod 2 = 0 Then
                    rw.DefaultCellStyle = styleAlternatingRowColor
                Else
                    rw.DefaultCellStyle = Nothing
                End If
            Next
        End With
        lblTransactionHistory.Text = "Transaction History - Item#: " & itm
    End Sub

    Private Sub UnselectDataGridViewRows(ByVal dgv As DataGridView)
        For Each rw As DataGridViewRow In dgv.Rows
            rw.Selected = False
        Next
    End Sub

    Private Sub ValidateTransactionDate()

        LoadCurrentPeriod()

        With systemperiod
            If Not cBusObj.ValidateTransactionDate(dtpTransactionDate.Value, .current_prd, .current_year) Then
                MsgBox("The month and year must be in the period " & .current_prd.ToString & "/" & .current_year.ToString & "." & vbCrLf & vbCrLf & _
                       "Enter a Transaction Effective Date that is within the period " & .current_prd.ToString & "/" & .current_year.ToString & ".", vbOKOnly, "Set Period Date")
                Exit Sub
            End If
        End With

    End Sub

    Private Function ValidateData()

        For Each rw As DataGridViewRow In dgvItems.Rows
            If rw.Cells(1).Value.ToString.Trim = "M" & txtItemSearch.Text.Trim Then
                Return False
            End If
        Next

        Return True

    End Function

    Private Function ConvertHTMLColor(ByVal HTMLColor As String)
        Dim RGB As Integer = Integer.Parse(HTMLColor.Replace("#", ""), System.Globalization.NumberStyles.HexNumber)
        Dim oColor As Color = Color.FromArgb(RGB)

        Return oColor
    End Function

    Private Sub Clear()
        dtItem.Clear()
        dtFactory001.Clear()
        dtWarehouse002.Clear()
        dgvWorkingArea.Rows.Clear()
        dtTransactionHistory.Clear()
        lblQtyCalculatorAlert.Visible = False
        txtMyQty.Text = ""
        txtTargetQty.Text = ""
        txtRemainderQty.Text = ""
        txtTargetQty.Enabled = True
        btnTransferRefresh.Enabled = False
        lblQtyCalculatorAlert.Visible = False
    End Sub

    Private Function RowsToRemove(ByVal dgv As DataGridView, ByVal itm As String) As String()
        Dim rws As Integer = 0
        Dim str As String = ""
        Dim sb As New StringBuilder
        Dim srws() As String

        Try
            With dgv
                For Each rw As DataGridViewRow In dgv.Rows
                    If rw.Cells(0).Value.ToString.Trim = itm Then
                        str = str & rw.Index.ToString & ","
                    End If
                Next
            End With
            str = str.Substring(0, str.Length - 1) 'remove last comma separator ...
            srws = str.Split(",")

            Return srws

        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Private Sub RemoveRows(ByVal rws As String())

        For i As Integer = rws.GetUpperBound(0) To 0 Step -1
            Dim deleterow As Integer = Convert.ToInt32(rws(i))
            dtWarehouse002.Rows.RemoveAt(deleterow)
            dtFactory001.Rows.RemoveAt(deleterow)
            dtItem.Rows.RemoveAt(deleterow)
        Next

    End Sub

    'Private Sub SplitBinsEqually(ItemNo As String, TargetLocTo As String, BinTo As String, TransactionQty As Integer, dgv As DataGridView)
    '    'currently unused.  purpose is to divide the bin quantities equally ...
    '    Dim bin_count As Integer = 0
    '    Dim equal_qty As Integer = 0
    '    Dim remainder_qty As Integer = 0
    '    Dim bins() As String = BinTo.Split
    '    Dim binqty("", 0) As Object
    '    'Start with splitting the qty here...

    '    With dgv

    '    End With
    'End Sub

#End Region

#Region "  REGION - Drag Drop to Working Area "

    Private dragBoxFromMouseDown As Rectangle
    Private rowIndexFromMouseDown As Integer
    Private rowIndexOfItemUnderMouseToDrop As Integer
    Private Sub Panel11_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Panel11.DragEnter
        e.Effect = DragDropEffects.Move
    End Sub
    Private Sub dgvWarehouse002_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
        e.Effect = DragDropEffects.Move
    End Sub
    Private Sub dgvWarehouse002_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles dgvWarehouse002.MouseDown, dgvFactory001.MouseDown
        Dim dgv As DataGridView = CType(sender, DataGridView)

        ' Get the index of the item the mouse is below.
        rowIndexFromMouseDown = dgv.HitTest(e.X, e.Y).RowIndex
        If rowIndexFromMouseDown = -1 Then Exit Sub
        If Not (My.Computer.Keyboard.ShiftKeyDown) Then
            If dgv.SelectedRows.Count <= 1 Then
                UnselectDataGridViewRows(dgv)
                dgv.Rows(rowIndexFromMouseDown).Selected = True
            End If
        End If

        If rowIndexFromMouseDown <> -1 Then
            ' Remember the point where the mouse down occurred. 
            ' The DragSize indicates the size that the mouse can move 
            ' before a drag event should be started.                
            Dim dragSize As Size = SystemInformation.DragSize

            Dim dropEffect As DragDropEffects = dgv.DoDragDrop(dgv.SelectedRows, DragDropEffects.Move) 'This for Multi Select Rows

        Else
            ' Reset the rectangle if the mouse is not over an item in the ListBox.
            dragBoxFromMouseDown = Rectangle.Empty
        End If

    End Sub
    Private Sub frmRepetitive_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter
        e.Effect = DragDropEffects.Move
    End Sub

    Private Function FillComoboxCell(dv As DataView, cel As DataGridViewCell) As DataGridViewComboBoxCell
        Dim cbo As DataGridViewComboBoxCell = cel

        For Each rw As DataRowView In dv
            cbo.Items.Add(rw("bin_no").ToString.Trim)
        Next

        Return cbo
    End Function
    Private Sub SetWorkingAreaCellColor(dgv As DataGridView, uidata As UICurrentData)
        With dgv
            If uidata.TranType = TransferType.Issue.ToString Then
                .Rows(.Rows.Count - 1).Cells(0).Style.BackColor = clrIssue
                .Rows(.Rows.Count - 1).Cells(1).Style.BackColor = clrIssue
                .Rows(.Rows.Count - 1).Cells("Direction").Style.BackColor = clrIssue
            ElseIf uidata.TranType = TransferType.Receipt.ToString Then
                .Rows(.Rows.Count - 1).Cells(0).Style.BackColor = clrReceipt
                .Rows(.Rows.Count - 1).Cells(1).Style.BackColor = clrReceipt
                .Rows(.Rows.Count - 1).Cells("Direction").Style.BackColor = clrReceipt
            ElseIf uidata.TranType = TransferType.Transfer.ToString Then
                .Rows(.Rows.Count - 1).Cells(0).Style.BackColor = clrTransfer
                .Rows(.Rows.Count - 1).Cells(1).Style.BackColor = clrTransfer
                .Rows(.Rows.Count - 1).Cells("Direction").Style.BackColor = clrTransfer
            End If
        End With

        UnselectDataGridViewRows(dgv)
    End Sub

    Private Sub PopulateTransactionType(uidata As UICurrentData)
        'Sub used to convert the raw uidata Object into a specific Transaction Type
        'then add it as an array to the datagridview
        Dim arr() As String = {"", "", "", "", "", "", "", 0, 0, 0, 0, 0}
        Dim cbowhsecell As New DataGridViewComboBoxCell
        Dim cbofctrycell As New DataGridViewComboBoxCell

        dvWareHouse = dtAllWarehouseBins.DefaultView
        dvWareHouse.RowFilter = "item_no = '" & uidata.ItemNo.Trim & "'"

        dvFactoryForCombo = dtFactoryForCombo.DefaultView

        Select Case uidata.TranType
            'Transfer uses From and To Values, populate all ...
            Case TransferType.Transfer.ToString
                arr(0) = uidata.TranType
                arr(1) = uidata.ItemNo
                arr(2) = uidata.TargetLocFrom
                arr(3) = uidata.BinFrom
                arr(4) = uidata.Direction
                arr(5) = uidata.TargetLocTo
                arr(6) = uidata.BinTo
                arr(7) = uidata.BinFromQtyOnHand
                arr(8) = uidata.BinFromTransQty
                arr(9) = uidata.BinToQtyOnHand  '
                arr(10) = uidata.BinToTransQty
                arr(11) = uidata.QtyToMove
                dgvWorkingArea.Rows.Add(arr)

                'For the ComboBoxCells, populate and show both From and To
                If uidata.TargetLocFrom = "002" Then
                    cbowhsecell = FillComoboxCell(dvWareHouse, dgvWorkingArea.Item("BinFirst", dgvWorkingArea.Rows.Count - 1))
                    cbowhsecell.Value = uidata.BinFrom
                    cbowhsecell.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                    cbowhsecell.ReadOnly = True
                Else
                    cbowhsecell = FillComoboxCell(dvFactoryForCombo, dgvWorkingArea.Item("BinFirst", dgvWorkingArea.Rows.Count - 1))
                    cbowhsecell.Value = uidata.BinFrom
                End If

                If uidata.TargetLocTo = "002" Then
                    cbofctrycell = FillComoboxCell(dvWareHouse, dgvWorkingArea.Item("BinSecond", dgvWorkingArea.Rows.Count - 1))
                    cbofctrycell.Value = uidata.BinTo
                Else
                    cbofctrycell = FillComoboxCell(dvFactoryForCombo, dgvWorkingArea.Item("BinSecond", dgvWorkingArea.Rows.Count - 1))
                    cbofctrycell.Value = uidata.BinTo
                End If

                btnTransferRefresh.Enabled = True
                txtTargetQty.Enabled = True
            Case TransferType.Issue.ToString
                'Issue uses the Left side of dgvWorkingArea Grid.  Fill ONLY the From Values, empty the To values
                arr(0) = uidata.TranType
                arr(1) = uidata.ItemNo
                arr(2) = uidata.TargetLocFrom
                arr(3) = uidata.BinFrom
                arr(4) = uidata.Direction
                arr(5) = ""
                arr(6) = ""
                arr(7) = uidata.BinFromQtyOnHand
                arr(8) = uidata.BinFromTransQty
                arr(9) = ""
                arr(10) = ""
                arr(11) = uidata.QtyToMove
                dgvWorkingArea.Rows.Add(arr)

                'For Issue, populate and show only the From ComboBoxCell, empty and hide the To ComboboxCells for Location and Bin
                If uidata.TargetLocFrom = "002" Then
                    cbowhsecell = FillComoboxCell(dvWareHouse, dgvWorkingArea.Item("BinFirst", dgvWorkingArea.Rows.Count - 1))
                    cbowhsecell.Value = uidata.BinFrom
                Else
                    cbowhsecell = FillComoboxCell(dvFactoryForCombo, dgvWorkingArea.Item("BinFirst", dgvWorkingArea.Rows.Count - 1))
                    cbowhsecell.Value = uidata.BinFrom
                End If

                'Line by Line: 
                'Add the cbofctrycell to the grid
                cbofctrycell = dgvWorkingArea.Item("BinSecond", dgvWorkingArea.Rows.Count - 1)
                'clear it of the items (it's populated at the beginning of this procedure) ... 
                cbofctrycell.Items.Clear()
                'set the display value to ""
                cbofctrycell.Value = ""
                'ComboBoxCells have a displayStyle, setting it to nothing essentially hides it
                cbofctrycell.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                'Finally, set it to ReadOnly, and the result is a cell the is empty and cannot be edited, just what we want!
                cbofctrycell.ReadOnly = True
                'Now, do the same for the ToLocation comboboxcell.  The only difference is we need to declare it (unlike the 
                'cbofctrycell, cbobintoloc hasn't been declared yet)
                'then we can add it to the grid, but everything else is identical to cbofctrycell ...
                Dim cbobintoloc As DataGridViewComboBoxCell = dgvWorkingArea.Item("TargetLocationTo", dgvWorkingArea.Rows.Count - 1)
                cbobintoloc.Items.Add("")
                cbobintoloc.DataGridView.Rows(dgvWorkingArea.Rows.Count - 1).Cells("TargetLocationTo").Value = ""
                cbobintoloc.Value = ""
                cbobintoloc.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                cbobintoloc.ReadOnly = True

                btnTransferRefresh.Enabled = False
                txtTargetQty.Enabled = False
            Case TransferType.Receipt.ToString
                'Receipt uses the Right side of dgvWorkingArea Grid.  Fill ONLY the To Values, empty the From values
                arr(0) = uidata.TranType
                arr(1) = uidata.ItemNo
                arr(2) = ""
                arr(3) = ""
                arr(4) = uidata.Direction
                arr(5) = uidata.TargetLocFrom 'This has to change from TargetLocationTo to TargetLocationFrom because the 'From' is the grid (001 or 002) it was dragged from and that is what displays
                arr(6) = uidata.BinFrom
                arr(7) = ""
                arr(8) = ""
                arr(9) = uidata.BinFromQtyOnHand  '
                arr(10) = uidata.BinFromTransQty
                arr(11) = 0   'uidata.QtyToMove
                dgvWorkingArea.Rows.Add(arr)

                'For Receipt, populate and show only the To ComboBoxCell, empty and hide the From ComboboxCells for Location and Bin
                If uidata.TargetLocFrom = "002" Then
                    cbofctrycell = FillComoboxCell(dvWareHouse, dgvWorkingArea.Item("BinSecond", dgvWorkingArea.Rows.Count - 1))
                    cbofctrycell.Value = uidata.BinFrom
                Else
                    cbofctrycell = FillComoboxCell(dvFactoryForCombo, dgvWorkingArea.Item("BinSecond", dgvWorkingArea.Rows.Count - 1))
                    cbofctrycell.Value = uidata.BinFrom
                End If

                'Line by Line: 
                'Add the cbofctrycell to the grid
                cbowhsecell = dgvWorkingArea.Item("BinFirst", dgvWorkingArea.Rows.Count - 1)
                'clear it of the items (it's populated at the beginning of this procedure) ... 
                cbowhsecell.Items.Clear()
                'set the display value to ""
                cbowhsecell.Value = ""
                'ComboBoxCells have a displayStyle, setting it to nothing essentially hides it
                cbowhsecell.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                'Finally, set it to ReadOnly, and the result is a cell the is empty and cannot be edited, just what we want!
                cbowhsecell.ReadOnly = True
                'Now, do the same for the ToLocation comboboxcell.  The only difference is we need to declare it (unlike the 
                'cbofctrycell, cbobintoloc hasn't been declared yet)
                'then we can add it to the grid, but everything else is identical to cbofctrycell ...
                Dim cbobinfromloc As DataGridViewComboBoxCell = dgvWorkingArea.Item("TargetLocationFrom", dgvWorkingArea.Rows.Count - 1)
                cbobinfromloc.Items.Add("")
                cbobinfromloc.DataGridView.Rows(dgvWorkingArea.Rows.Count - 1).Cells("TargetLocationFrom").Value = ""
                cbobinfromloc.Value = ""
                cbobinfromloc.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                cbobinfromloc.ReadOnly = True

                btnTransferRefresh.Enabled = False
                txtTargetQty.Enabled = False
        End Select

        SetWorkingAreaCellColor(dgvWorkingArea, uidata)
        dgvComboLocFrom.ReadOnly = True
        dgvComboLocFrom.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing

    End Sub

    Private Sub PopulateUIDataWorkArea(dtTo As DataTable, dtFrom As DataTable)
        'This Sub arranges the values needed for any of the three transaction types, Transfer, Issue or Receipt,   
        'and puts them in the object cUIDataWrkArea.  The specific Transaction type is finalized in the next function
        'and prepared as a DataGridViewRow

        Dim countFrom As Integer = 0
        Dim countTo As Integer = 0
        Dim rowcount As Integer = 0
        Dim fromGridLoc As String = ""
        Dim qtyToMoveLoc As String
        Dim cUIDataWrkArea As New UICurrentData
        Dim qtyToMoveFromTableName As String
        Dim arr() As String = {"", "", "", "", "", "", "", 0, 0, 0, 0, 0}

        'Need to determine the row count and take the greater of the From and To so we handle all the data being moved ... 

        countFrom = dtFrom.Rows.Count
        countTo = dtTo.Rows.Count

        If countFrom >= countTo Then
            rowcount = countFrom
            qtyToMoveLoc = dtFrom(0)("targetloc_from")
            qtyToMoveFromTableName = dtFrom.TableName
        Else
            rowcount = countTo
            qtyToMoveLoc = dtTo(0)("targetloc_to")
            qtyToMoveFromTableName = dtTo.TableName
        End If

        With cUIDataWrkArea
            .TranType = dtFrom(0)("trantype")
            .TargetLocFrom = dtFrom(0)("targetloc_from").ToString.Trim
            .TargetLocTo = dtTo(0)("targetloc_to")   'IIf(.TargetLocFrom = "001", "002", "001")
            .ItemNo = dtFrom(0)("itemno")

            'WHY is this for loop reversed here: Because the loop adds from the datatable from 
            '    from the lowest number bin to the highest, that puts the list in DESC order when
            '    it is used in the grid, and because the Bins are really comoboboxes, they can't 
            '    be sorted reasonably.  So, reverse it here and it's correct in the dgvWorkingArea Grid!
            For i As Integer = 0 To rowcount - 1  'rowcount - 1 To 0 Step -1
                'Set the bins, from and to.  
                If dtFrom(i) Is Nothing Then .BinFrom = dtFrom(0)("binfrom") Else .BinFrom = dtFrom(i)("binfrom")
                If dtTo(i) Is Nothing Then If dtTo(0) Is Nothing Then .BinTo = "" Else .BinTo = "" Else .BinTo = dtTo(i)("binto")
                'NO GOOD - If dtTo(i) Is Nothing Then If dtTo(0) Is Nothing Then .BinTo = "" Else .BinTo = dtTo(0)("binto") Else .BinTo = dtTo(i)("binto")
                'NO GOOD - If dtFrom(i) Is Nothing Then .BinFrom = "" Else .BinFrom = dtFrom(i)("binfrom")

                'NOTE on Quantities populated here: because there may be multiple rows in one datatable but only one row in the other (i.e. dtFrom has 3 rows but dtTo has only 1 row),
                'we are checking for the next records first (i.e. If dtFrom(i) is nothing) and if it is Nothing, then we
                'fall back to the first record, dtFrom(0) ...
                'These are default values directly from the grid
                If dtFrom(i) Is Nothing Then cUIDataWrkArea.BinFromQtyOnHand = dtFrom(0)("bin_from_qty_on_hand") Else cUIDataWrkArea.BinFromQtyOnHand = dtFrom(i)("bin_from_qty_on_hand")
                If dtFrom(i) Is Nothing Then cUIDataWrkArea.BinFromTransQty = dtFrom(0)("bin_from_transqty") Else cUIDataWrkArea.BinFromTransQty = dtFrom(i)("bin_from_transqty")
                If dtTo(i) Is Nothing Then If dtTo(0) Is Nothing Then cUIDataWrkArea.BinToQtyOnHand = 0 Else cUIDataWrkArea.BinToQtyOnHand = dtTo(0)("bin_to_qty_on_hand") Else cUIDataWrkArea.BinToQtyOnHand = dtTo(i)("bin_to_qty_on_hand")
                If dtTo(i) Is Nothing Then If dtTo(0) Is Nothing Then cUIDataWrkArea.BinToTransQty = 0 Else cUIDataWrkArea.BinToTransQty = dtTo(0)("bin_to_transqty") Else cUIDataWrkArea.BinToTransQty = dtTo(i)("bin_to_transqty")

                'The .QtyToMove value is used to set a default quantity when we need to know what the previous Transfer Qty was; 
                'This is done by checking the IMREPLOG_MAS table using GetQtytoMove function and returning what the last trans Qty was.  
                'NOTE: GetQtyToMove only applies to the dtTo table, since the dtFrom table is assumed to have qty in it to be able to move.  
                If qtyToMoveFromTableName = "dtFrom" Then cUIDataWrkArea.QtyToMove = .BinFromTransQty Else cUIDataWrkArea.QtyToMove = GetQtytoMove(cUIDataWrkArea.TranType, cUIDataWrkArea.BinTo, cUIDataWrkArea.ItemNo) '.BinToTransQty

                Select Case cUIDataWrkArea.TranType
                    Case TransferType.Transfer.ToString
                        .Direction = sTransactionTRAN     'sTransactionIN & sTransactionOUT
                    Case TransferType.Issue.ToString
                        .Direction = sTransactionOUT
                    Case TransferType.Receipt.ToString
                        .Direction = sTransactionIN
                End Select

                PopulateTransactionType(cUIDataWrkArea)

            Next

        End With

    End Sub

    Private Sub Panel11_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Panel11.DragDrop
        Dim x As Integer = Me.PointToClient(New Point(e.X, e.Y)).X  '575
        Dim y As Integer = Me.PointToClient(New Point(e.X, e.Y)).Y  '306
        Dim parentCtrl As Panel = CType(sender, Panel)
        Dim picbox As String = getPicbox(x, y, parentCtrl)
        Dim itemCount As Integer = 0 'itemCount is the Qty column (or total count of the items) in LocFrom
        Dim countFrom As Integer = 0
        Dim countTo As Integer = 0
        Dim rowcount As Integer = 0
        Dim sSQL As String = ""
        Dim arr() As String = {"", "", "", "", "", "", "", 0, 0, 0, 0, 0}

        'STEP 1 - dtFrom:  Convert the Dragged data to a datatable so we can work with it more reasonably ...
        Dim dtFrom As DataTable = LoadWorkingAreaDataTable() 'This just creates the empty table structure to populate
        dtFrom.TableName = "dtFrom"

        'this creates the empty rows collection object for populating and adding to the dtFrom datatable
        Dim rws As DataGridViewSelectedRowCollection = Nothing

        'STEP 2 - Now it's time to fill the UIData Object and the empty rws collection with data ...
        With uidata
            'TranType comes from determining which PictureBox was drag/dropped on ...
            .TranType = getPicbox(x, y, parentCtrl)

            'The data in the dragged 'format', which is a datagridview row, is put into the rws collection.  
            'Multiple rows can be selected ... 
            If e.Effect = DragDropEffects.Move Then
                For Each format As String In e.Data.GetFormats()
                    rws = e.Data.GetData(format)
                Next

                Dim dt As DataTable = GetDataTableFromDataRowsCollection(rws)
                Dim dv As DataView = dt.DefaultView
                dv.Sort = "bin ASC"
                dt = dv.ToTable

                'these values are common to all rows, i.e. From and To (001, 002) and the Item_No.  
                .TargetLocFrom = dt(0)("loc").ToString.Trim
                .TargetLocTo = IIf(.TargetLocFrom = "001", "002", "001")
                .ItemNo = dt(0)("itemno").ToString.Trim

                Try
                    For Each rw As DataRow In dt.Rows
                        Dim dr As DataRow = dtFrom.NewRow
                        dr(0) = uidata.TranType
                        dr(1) = rw("ItemNo").ToString.Trim
                        dr(2) = rw("bin").ToString.Trim
                        dr(3) = rw("loc").ToString.Trim
                        dr(4) = Convert.ToInt32(Convert.ToDecimal(rw("qty")))
                        'If .TranType = TransferType.Receipt.ToString Then dr(4) = 0 Else dr(4) = Convert.ToInt32(rw("qty"))
                        dr(5) = Convert.ToInt32(Convert.ToDecimal(rw("qty")))
                        dr(6) = ""
                        dr(7) = ""
                        dr(8) = 0
                        dr(9) = 0
                        dtFrom.Rows.Add(dr)
                        itemCount = itemCount + dr(4)
                    Next
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

            End If

            'STEP 4 - Create & Populate dtTo ...
            Dim dtTo As New DataTable
            'Deterimine where the ItemNo is stored, if it's in 001 or 002. 

            If .TargetLocFrom = "002" Then
                .TargetLocTo = "001"
            Else
                'GetCurrentBins function checks the last BIN transaction from the IMBINTRX_SQL to return to the warehouse the qty taken out ...
                Dim dtloc As DataTable = cBusObj.GetItemLocation(.ItemNo, cn)
                .TargetLocTo = dtloc(0)(0)
            End If

            dtTo = cBusObj.GetCurrentBins(.ItemNo, itemCount, .TargetLocTo, cn)

            If dtTo.Rows.Count = 0 Then
                'check to be sure we have a default bin to Transfer To
                If cboDefaultFactoryBin.Text = "" Then
                    fFactoryBinSelect.ShowDialog()
                    'check to make sure we did select a Bin, if not exit ...
                    If cboDefaultFactoryBin.Text = "" Then
                        MsgBox("Select a Factory Bin from 'Set Factory Bin' drop down below Factory Bins Grid to continue.", MsgBoxStyle.OkOnly, "Factory Bin Not Set")
                        Exit Sub
                    End If
                End If
                'now we know we have a bin, so let's create the dtTo table manually
                dtTo = GetDTToManually(.TranType, .ItemNo, .TargetLocTo, cboDefaultFactoryBin.Text.Trim)
            End If
            PopulateUIDataWorkArea(dtTo, dtFrom)
        End With

    End Sub
    Private Function GetDTToManually(tranType As String, itemNo As String, targetLocTo As String, binTo As String) As DataTable
        Dim dt As New DataTable
        With dt
            .Columns.Add("trantype")
            .Columns.Add("itemno")
            .Columns.Add("binfrom")
            .Columns.Add("targetloc_from")
            .Columns.Add("bin_from_qty_on_hand")
            .Columns.Add("bin_from_transqty")
            .Columns.Add("binto")
            .Columns.Add("targetloc_to")
            .Columns.Add("bin_to_qty_on_hand")
            .Columns.Add("bin_to_transqty")
        End With
        Dim ar() As String = {tranType, itemNo, "", "", 0, 0, binTo, targetLocTo, 0, 0}
        dt.Rows.Add(ar)

        Return dt
    End Function
    Private Function GetQtytoMove(tran_type As String, bin_from As String, item_no As String) As Integer

        Dim sSQL As String = "Select top 1 qty_to_move " & vbCrLf & _
                             "from IMREPLOG_MAS " & vbCrLf & _
                             "where tran_type = '" & tran_type & "' " & vbCrLf & _
                             "and bin_from = '" & bin_from & "' " & vbCrLf & _
                             "and item_no = '" & item_no & "' " & vbCrLf & _
                             "order by A4GLIdentity DESC"
        Dim QtyToMove As Integer = Convert.ToInt32(cBusObj.ExecuteSQLScalar(sSQL, cn))
        If QtyToMove = Nothing Then QtyToMove = 0

        Return QtyToMove

    End Function


    Private Function ValidateWorkingAreaData(ValidationType As Integer, TransactionType As String, bin As String, loc As String, Optional Qty As Object = 0, Optional ItemNo As String = "") As Boolean
        Dim isValid As Boolean = True

        Select Case ValidationType
            'Quantity must be validated on Transfers ...
            Case ValidateType.ValidQuantity
                If Convert.ToDecimal(Qty) = 0 And TransactionType = TransferType.Transfer.ToString Then
                    MsgBox("- Zero Quantity -" & vbCrLf & vbCrLf & "Transfers must have quantity in the 'From' bin.  " & vbCrLf & vbCrLf & _
                           "'From' Bin " & bin.Trim & " at location " & loc.Trim & " has a quantity of " & Convert.ToDecimal(Qty) & vbCrLf & vbCrLf & "Please drag from a bin and location that has quantity", MsgBoxStyle.OkOnly, "'From' bin has 0 quantity")
                    isValid = False
                End If
            Case ValidateType.Duplicated
                Dim dgv As DataGridView = CType(dgvWorkingArea, DataGridView)
                For Each rw As DataGridViewRow In dgv.Rows
                    Dim item_no As String = rw.Cells("ItemNo").Value.ToString.Trim
                    Dim binfrom As String = rw.Cells("BinFirst").Value.ToString.Trim
                    Dim locfrom As String = rw.Cells("TargetLocationFrom").Value.ToString.Trim
                    If item_no = ItemNo AndAlso binfrom = bin AndAlso locfrom = loc Then
                        MsgBox("- Duplicate Entry - " & vbCrLf & vbCrLf & "Duplicate transactions are not allowed.  " & vbCrLf & vbCrLf & _
                                                   "Item No " & item_no & " at  Bin " & bin.Trim & " and location " & loc.Trim & " is already in the transaction list. ")

                        isValid = False
                    End If
                Next

        End Select

        Return isValid
    End Function

    Private Function GetWorkingAreaData(TargetLocTo As String, picbox As String, ReturnType As String, Optional targettype As String = "", Optional bin As String = "", Optional ItemNo As String = "") As String

        Dim sReturn As String = ""
        Try
            Select Case ReturnType

                Case ReturnDataType.BinTo.ToString
                    'For the return bins (BinTo) on a transfer, we need to determin if a bin is already assigned to the item at the location
                    'If the 'To' Location is the Warehouse, we need to build a string that will contain all the bins for that item ...
                    If TargetLocTo = "002" Then
                        Dim dr() As DataRow = dtAllWarehouseBins.Select("item_no = '" & ItemNo & "'")
                        Dim dt As DataTable = dr.CopyToDataTable
                        If dt Is Nothing OrElse dt(0)("bin_no").ToString.Trim = "" Then
                            sReturn = ""
                        Else
                            For Each rw As DataRow In dt.Rows
                                If sReturn = "" Then
                                    sReturn = rw("bin_no").ToString.Trim
                                Else
                                    sReturn = sReturn & ", " & rw("bin_no").ToString.Trim
                                End If
                            Next
                        End If
                    Else
                        'If the 'To' location is the Factory, there is only ever one bin where it is sent ...
                        Dim dr() As DataRow = dtAllFactoryBins.Select("item_no = '" & ItemNo & "'")
                        Dim dt As DataTable = dr.CopyToDataTable
                        If dt Is Nothing OrElse dt(0)("bin_no").ToString.Trim = "" Then
                            sReturn = ""
                        Else
                            sReturn = dt(0)("bin_no").ToString.Trim
                        End If
                    End If

                Case ReturnDataType.Direction.ToString
                    If picbox = TransferType.Transfer.ToString Or picbox = TransferType.Issue.ToString Then
                        sReturn = "   " & sTransactionOUT & "   "
                    Else
                        sReturn = "   " & sTransactionOUT & "   "
                    End If

            End Select

        Catch ex As Exception

        End Try

        Return (sReturn)

    End Function
    Private Function getPicbox(x As Integer, y As Integer, parentCtl As Control) As String
        Dim picbox As PictureBox
        Dim picname As String = ""

        If parentCtl.HasChildren Then
            For Each ctrl As Control In parentCtl.Controls
                If TypeOf (ctrl) Is PictureBox Then
                    picbox = CType(ctrl, PictureBox)

                    Dim picbox_x As Integer = picbox.Location.X + Panel10.Width + SplitContainer3.Location.X
                    Dim picbox_y As Integer = picbox.Location.Y + Panel3.Height + SplitContainer3.Location.Y

                    Dim picboxPoint As Point = New Point(picbox_x, picbox_y)

                    If ((x >= picboxPoint.X) _
                                       AndAlso ((x _
                                       <= (picboxPoint.X + picbox.Width)) _
                                       AndAlso ((y >= picboxPoint.Y) _
                                       AndAlso (y _
                                       <= (picboxPoint.Y + picbox.Height))))) Then

                        picname = picbox.Tag
                        Exit For
                    End If

                End If
            Next
        End If



        Return picname
    End Function

#End Region

#Region "  REGION - QTY Calculator on Transfers"

    Private Sub txtDGVQuantity_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtDGVQuantity.KeyPress
        'txtDGVQuantity is a TextBox derived from the WorkingArea DGV, it doesn't exist until is the DGV is edited ...
        Dim txt As TextBox = DirectCast(txtDGVQuantity, TextBox)
        Dim sval As String = "0"
        Dim qty As Integer = 0
        If Char.IsNumber(e.KeyChar) Then
            sval = txtDGVQuantity.Text
        End If

        With Timer3
            .Interval = 500
            .Enabled = True
        End With
    End Sub

    'Private Sub txtDGVQuantity_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtDGVQuantity.KeyUp
    '    If e.KeyCode = Keys.Enter Then
    '        iDoneCounter += 1
    '        With Timer5
    '            .Interval = 2000
    '            .Enabled = True
    '        End With
    '        If iDoneCounter = 2 Then
    '            Dim dgv As DataGridView = CType(sender, DataGridView)
    '            With dgv
    '                If .Columns(.CurrentCell.ColumnIndex).Name = "QtyToMove" Then
    '                    iDoneCounter = 0
    '                    WorkAreaDone()
    '                End If
    '            End With
    '        End If
    '    End If
    'End Sub


    Private Sub txtDGVQuantity_TextChanged(sender As Object, e As System.EventArgs) Handles txtDGVQuantity.TextChanged
        Dim txt As TextBox = DirectCast(txtDGVQuantity, TextBox)
        Dim sval As String = "0"
        Dim myqty As Integer = 0
        Dim dgv As DataGridView = CType(dgvWorkingArea, DataGridView)
        With dgv
            If .Rows(.CurrentCell.RowIndex).Cells("TranType").Value.ToString.Trim <> TransferType.Transfer.ToString Then Exit Sub
            If IsNumeric(txt.Text) Then

                sval = Convert.ToInt32(txt.Text)
                myqty = Convert.ToInt32(sval)
                CalculateQuantityToMove(myqty)
            End If
        End With
    End Sub

    Private Sub SetCalculatorLabel(myqty As Integer, Target As Integer)
        If txtTargetQty.Text.Trim = "" Then
            lblQtyCalculatorAlert.Visible = False
            txtMyQty.Text = ""
            txtRemainderQty.Text = ""
            Exit Sub
        End If


        Dim lbl As Label = CType(lblQtyCalculatorAlert, Label)
        Dim dif As Integer = myqty - Target
        With lbl
            Select Case dif
                Case 0
                    .Text = "OK"
                    .ForeColor = Color.Green
                Case Is < 0
                    .Text = "Under"
                    .ForeColor = Color.Blue
                Case Is > 0
                    .Text = "Over"
                    .ForeColor = Color.Crimson
            End Select
            .Visible = True
        End With
    End Sub

    Private Sub txtTargetQty_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtTargetQty.TextChanged
        Dim txt As TextBox = CType(sender, TextBox)
        Dim myqty As Integer = 0
        Dim target As Integer = 0
        If IsNumeric(txtTargetQty.Text) Then target = Convert.ToInt32(txtTargetQty.Text)
        If txtMyQty.Text = "" OrElse txtMyQty.Text = "0" Then
            Dim dgv As DataGridView = CType(dgvWorkingArea, DataGridView)
            myqty = CalculateDataGridViewQty(dgv)
        ElseIf IsNumeric(txtMyQty.Text) Then
            myqty = Convert.ToInt32(txtMyQty.Text)
        Else
            myqty = 0
        End If

        txtMyQty.Text = myqty
        txtRemainderQty.Text = myqty - target

        If txtRemainderQty.Text <> "" OrElse txtRemainderQty.Text <> "0" Then
            SetCalculatorLabel(myqty, target)
        End If
    End Sub

    Private Sub CalculateQuantityToMove(myqty As Integer)
        Dim dgv As DataGridView = CType(dgvWorkingArea, DataGridView)
        Dim targetQty As String = IIf(txtTargetQty.Text = "", "0", txtTargetQty.Text)
        Dim rwidx As Integer = dgv.CurrentCell.RowIndex
        Try

            Dim target As Integer = Convert.ToInt32(targetQty)
            Dim qtydgv As Integer = CalculateDataGridViewQty(dgv, rwidx)

            myqty = myqty + qtydgv
            txtMyQty.Text = myqty
            txtRemainderQty.Text = myqty - target

            SetCalculatorLabel(myqty, target)

        Catch ex As Exception

        End Try

    End Sub

    Private Function CalculateDataGridViewQty(dgv As DataGridView, rwidx As Integer) As Integer
        Dim qtydgv As Integer = 0
        Dim itemno As String = dgv.Rows(rwidx).Cells("ItemNo").Value.ToString.Trim

        For Each rw As DataGridViewRow In dgv.Rows
            If rw.Index <> rwidx AndAlso rw.Cells("QtyToMove").Value <> "" Then
                If rw.Cells("ItemNo").Value.ToString.Trim = itemno AndAlso rw.Cells("TranType").Value.ToString = TransferType.Transfer.ToString Then
                    qtydgv = qtydgv + Convert.ToInt32(rw.Cells("QtyToMove").Value)
                End If

            End If


        Next
        Return qtydgv
    End Function

    Private Function ValidataDataGridViewQty(dgv As DataGridView, rwidx As Integer) As Boolean
        Dim validated As Boolean
        Dim qtymove As Integer = 0
        Dim qtybinfrom As Integer = 0
        With dgv
            .EndEdit()
            qtymove = IIf(.Rows(rwidx).Cells("QtyToMove").Value = "" OrElse .Rows(rwidx).Cells("QtyToMove").Value = "0", 0, Convert.ToInt32(.Rows(rwidx).Cells("QtyToMove").Value))
            qtybinfrom = IIf(.Rows(rwidx).Cells("BinFromQtyOnHand").Value = "" OrElse .Rows(rwidx).Cells("BinFromQtyOnHand").Value = "0", 0, Convert.ToInt32(.Rows(rwidx).Cells("BinFromQtyOnHand").Value))
            If qtymove > qtybinfrom Then
                validated = False
            Else
                validated = True
            End If
        End With

        Return validated
    End Function

    Private Function CalculateDataGridViewQty(dgv As DataGridView) As Integer
        Dim qtydgv As Integer = 0
        Try
            Dim itemno As String = dgv.Rows(dgv.CurrentCell.RowIndex).Cells("ItemNo").Value.ToString.Trim

            For Each rw As DataGridViewRow In dgv.Rows
                If rw.Cells("QtyToMove").Value <> "" Then
                    If rw.Cells("ItemNo").Value.ToString.Trim = itemno _
                        AndAlso rw.Cells("TranType").Value.ToString.Trim = TransferType.Transfer.ToString Then
                        qtydgv = qtydgv + Convert.ToInt32(rw.Cells("QtyToMove").Value)
                    End If
                End If
            Next
        Catch ex As Exception

        End Try

        Return qtydgv
    End Function

    Private Function DoQtyCalculation()
        Dim iTotal As Integer = 0
        Dim dgv As DataGridView = CType(dgvWorkingArea, DataGridView)

        With dgv
            For Each rw As DataGridViewRow In dgv.Rows
                iTotal = iTotal
            Next
        End With

        Return iTotal

    End Function

#End Region

#Region "  REGION -  Edit Working Area DGV and Related Control Events "
    'TODO - Remove is not used
    'Private Sub MyDataGridViewInitializationMethod()
    '    AddHandler dgvWorkingArea.EditingControlShowing, AddressOf Me.dgvWorkingArea_EditingControlShowing
    'End Sub
    'TODO - Start here on Cleanup ...
    Private Sub dgvWorkingArea_EditingControlShowing(sender As Object, e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvWorkingArea.EditingControlShowing
        'EditingControlShowing Event handles filling the comboboxes with the correct bins and Locations based on the drag/drop selection.  It also 
        'gets the textbox value needed to do the QTY Calculator on Transfers.  

        Dim dgv As DataGridView = CType(sender, DataGridView)
        dgvFromEditControlShowing = dgv 'This is used for ContextMenuStrip1 when it is called from EditControlShowing
        Dim combo As New ComboBox
        Dim itemNo As String
        'Dim pnl As Panel
        Dim cel As DataGridViewCell = dgv.CurrentCell
        Dim col As Integer
        Dim row As Integer
        Dim colname As String = ""

        e.Control.ContextMenuStrip = ContextMenuStrip1

        'Step 1 - If we're on a Combobox cell, it has to be BinFirst or BinSecond, otherwise 
        '         we're on a textbox cell which we'll use with QTY Calculator ...
        If TypeOf e.Control Is System.Windows.Forms.DataGridViewComboBoxEditingControl Then
            combo = CType(e.Control, ComboBox)
            'pnl = combo.Parent ' this is the panel around the cell
            dgv = combo.Parent.Parent 'combo is comboboxcell, parent1 is the panel around the cell, and parent2 is the datagridview
            cel = dgv.CurrentCell
            col = cel.ColumnIndex
            row = cel.RowIndex
            colname = dgv.Columns(col).Name
            If Not (colname = "BinFirst" Or colname = "BinSecond") Then Exit Sub
        Else
            Dim txt = CType(e.Control, TextBox)
            'pnl = txt.Parent ' this is the panel around the cell
            dgv = txt.Parent.Parent 'combo is comboboxcell, parent1 is the panel around the cell, and parent2 is the datagridview
            cel = dgv.CurrentCell
            col = cel.ColumnIndex
            row = cel.RowIndex
            colname = dgv.Columns(col).Name
        End If

        If Not (TypeOf e.Control Is System.Windows.Forms.DataGridViewComboBoxEditingControl) Then
            If colname = "QtyToMove" Then
                txtDGVQuantity = CType(e.Control, TextBox) 'txtDGVQuantity only exists when assigned the DGV Text Box Control in Edit Control Showing Event
                dgv.CommitEdit(DataGridViewDataErrorContexts.CurrentCellChange)
                AddHandler e.Control.KeyPress, AddressOf txtDGVQuantity_KeyPress
                'AddHandler e.Control.KeyUp, AddressOf txtDGVQuantity_KeyUp
                Exit Sub
            Else
                Exit Sub
            End If
        End If

        'Step 2 - If we have BinFirst or BinSecond, fill the comoboboxes with bins to match the 
        '       - TargetLocationFrom (for Binfirst) or TargetLocationTo (for BinSecond) and 
        '       - and fill with available bins from the Warehouse or Factory

        'busy work, get ItemNo, clear existing Items 
        itemNo = dgv.Rows(row).Cells("ItemNo").Value.ToString.Trim
        combo.Items.Clear()

        Dim loc As String 'loc will hold the Location that relates to BinFirst or BinSecond
        Dim cbowhsecell As New DataGridViewComboBoxCell
        Dim cbofctrycell As New DataGridViewComboBoxCell

        'Populate the DataViews with data ...
        dvWareHouse = dtAllWarehouseBins.DefaultView
        dvWareHouse.RowFilter = "item_no = '" & itemNo & "'"
        dvFactoryForCombo = dtFactoryForCombo.DefaultView

        Try
            Select Case colname
                Case "BinFirst"
                    loc = dgv.Rows(row).Cells("TargetLocationFrom").Value.ToString.Trim
                    If loc = "002" Then
                        For Each rw As DataRowView In dvWareHouse
                            combo.Items.Add(rw("bin_no").ToString.Trim)
                        Next
                        cbowhsecell = FillComoboxCell(dvWareHouse, dgvWorkingArea.Item("BinFirst", row))
                        cbowhsecell.Value = uidata.BinFrom
                    Else
                        For Each rw As DataRowView In dvFactoryForCombo
                            combo.Items.Add(rw("bin_no").ToString.Trim)
                        Next
                        cbowhsecell = FillComoboxCell(dvFactoryForCombo, dgvWorkingArea.Item("BinFirst", row))
                        cbowhsecell.Value = uidata.BinFrom
                    End If
                Case "BinSecond"
                    loc = dgv.Rows(row).Cells("TargetLocationTo").Value.ToString.Trim
                    If loc = "002" Then
                        For Each rw As DataRowView In dvWareHouse
                            combo.Items.Add(rw("bin_no").ToString.Trim)
                        Next
                        cbofctrycell = FillComoboxCell(dvWareHouse, dgvWorkingArea.Item("BinSecond", row))
                        cbofctrycell.Value = uidata.BinTo
                    Else
                        For Each rw As DataRowView In dvFactoryForCombo
                            combo.Items.Add(rw("bin_no").ToString.Trim)
                        Next
                        cbofctrycell = FillComoboxCell(dvFactoryForCombo, dgvWorkingArea.Item("BinSecond", row))
                        cbofctrycell.Value = uidata.BinTo
                    End If
            End Select

            e.CellStyle.BackColor = dgv.DefaultCellStyle.BackColor

        Catch ex As Exception

        End Try

    End Sub

    Private Sub ComboBox1_DropDownClosed(sender As Object, e As System.EventArgs) Handles comboBinSecond.DropDownClosed, comboBinFirst.DropDownClosed
        Dim dgv As DataGridView = CType(dgvWorkingArea, DataGridView)
        Dim cbo As ComboBox = CType(sender, ComboBox)

        'The dgvCellEndEdit Event needs to be called to set the combobox.SelectedItem value, need to create the event parameters here...
        Dim eventargs As New DataGridViewCellEventArgs(dgv.CurrentCell.ColumnIndex, dgv.CurrentRow.Index)

        dgvWorkingArea_CellEndEdit(dgv, eventargs)
        dgv.CurrentCell = dgv.Rows(dgv.CurrentRow.Index).Cells("direction")
    End Sub

    Private Sub combobox_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles comboLocationFrom.SelectedValueChanged, comboLocationTo.SelectedValueChanged, _
                                                                                                   comboBinFirst.SelectedValueChanged, comboBinSecond.SelectedValueChanged
        Dim cbo As ComboBox = CType(sender, ComboBox)
        Dim cbovalue As String = cbo.SelectedItem
        Dim dgv As DataGridView = CType(dgvWorkingArea, DataGridView)

        Select Case cbo.Name
            Case "comboLocationTo"
                If uidata.ComboLocationToValue = cbovalue Then

                    Exit Sub
                Else
                    dgv.Rows(dgv.CurrentRow.Index).Cells("BinSecond").Value = "" 'If we are changing the Location, we need to eliminate the Bin as well, set to ""
                    uidata.ComboLocationToValue = cbovalue
                    'For Each rw As DataGridViewRow In dgv.Rows
                    '    rw.Cells("TargetLocationTo").Value = uidata.ComboLocationToValue
                    '    dgvComboLocTo.Items.Add(uidata.ComboLocationToValue)
                    '    'dgvWorkingArea.Rows(dgvWorkingArea.Rows.Count - 1).Cells("TargetLocationTo").Value = uidata.ComboLocationToValue
                    '    rw.Cells("TargetLocationTo").Value = uidata.ComboLocationToValue
                    '    rw.Cells("BinSecond").Value = ""
                    'Next
                End If
            Case "comboLocationFrom"
                If uidata.ComboLocationFromValue = cbovalue Then
                    comboLocationFrom.Tag = cbovalue
                    'TextBox1.Text = cbovalue
                    Exit Sub
                Else
                    dgv.Rows(dgv.CurrentRow.Index).Cells("BinFirst").Value = "" 'this empties the bin associated with the location if it changes, otherwise bins will show as available at location where they don't exist
                    uidata.ComboLocationFromValue = cbovalue 'now, save the combobox value to the class
                    comboLocationFrom.Tag = cbovalue
                    'TextBox1.Text = cbovalue
                End If

            Case "comboBinFirst"
                If uidata.ComboBinFirstValue = cbovalue Then
                    'TextBox5.Text = cbovalue
                    Exit Sub
                Else
                    'dgv.Rows(dgv.CurrentRow.Index).Cells("BinFirst").Value = ""
                    uidata.ComboBinFirstValue = cbovalue
                    'comboBinFirst.Tag = ""
                    'TextBox5.Text = cbovalue
                End If

            Case "comboBinSecond"

                uidata.ComboBinSecondValue = cbovalue


        End Select
        cbo = Nothing
    End Sub

    Private Sub dgvWorkingArea_CellEndEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWorkingArea.CellEndEdit
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim new_bin_no As String = ""
        Dim item_no As String = ""
        Dim loc As String = ""
        Dim qty_on_hand As Integer = 0
        Dim counter As Integer = 0
        Dim colname As String = ""
        Dim sComboSource As String = ""
        Dim combo As New ComboBox

        Try
            Dim cel As DataGridViewCell = dgv.Rows(e.RowIndex).Cells(e.ColumnIndex)
            Dim col As Integer = cel.ColumnIndex
            colname = dgv.Columns(col).Name

            Select Case colname
                'This procedure is intended as a way to capture a manually entered value into the combobox which is then saved in the datatable for future use ...
                Case "BinFirst"
                    'sComboSource = "002" 'comboBinFirst.Tag

                    combo = CType(comboBinFirst, ComboBox)
                    Dim txt As String = ""
                    If Not combo.Text = "" Then txt = combo.Text Else If Not combo.SelectedItem = "" Then txt = combo.SelectedItem Else txt = ""
                    If txt = "" Then
                        'Exit Try

                    Else
                        new_bin_no = txt
                        For Each itm As String In combo.Items
                            If itm.Trim = new_bin_no Then
                                counter = 1
                                Exit For
                            End If
                        Next
                    End If
                    If counter = 0 Then
                        combo.Items.Add(new_bin_no.Trim)

                        With dgv
                            item_no = .CurrentRow.Cells("ItemNo").Value.ToString.Trim
                            loc = "00" & CStr(Locations.Warehouse)
                            qty_on_hand = 0
                        End With

                        If sComboSource = Locations.Warehouse.ToString Then
                            With dtAllWarehouseBins
                                Dim rw As DataRow = dtAllWarehouseBins.NewRow
                                rw("item_no") = item_no
                                rw("loc") = loc
                                rw("bin_no") = new_bin_no
                                rw("qty_on_hand") = 0
                                .Rows.Add(rw)
                            End With
                        Else
                            With dtAllFactoryBins
                                Dim rw As DataRow = dtAllWarehouseBins.NewRow
                                rw("item_no") = item_no
                                rw("loc") = loc
                                rw("bin_no") = new_bin_no
                                rw("qty_on_hand") = 0
                                .Rows.Add(rw)
                            End With
                        End If

                        dgvComboBinFirst.Items.Add(new_bin_no)
                        dgvComboBinFirst.Items(0).ToString()
                    Else
                        Dim item As String = txt
                        counter = 0
                        With dgvComboBinFirst
                            For Each itm As String In .Items
                                If itm.ToString.Trim = item Then
                                    counter = 1
                                    Exit For
                                End If
                            Next
                        End With
                        If counter = 0 Then
                            dgvComboBinFirst.Items.Add(item)
                        End If

                        dgv.CurrentRow.Cells("BinFirst").Value = new_bin_no
                    End If

                Case "BinSecond"
                    sComboSource = comboBinSecond.Tag
                    combo = CType(comboBinSecond, ComboBox)
                    Dim txt As String = ""
                    If Not combo.Text = "" Then txt = combo.Text Else If Not combo.SelectedItem = "" Then txt = combo.SelectedItem Else txt = ""
                    If txt = "" Then
                        'Exit Try
                    Else
                        new_bin_no = txt
                        For Each itm As String In combo.Items
                            If itm = new_bin_no Then
                                counter = 1
                            End If
                        Next
                    End If
                    If counter = 0 Then
                        combo.Items.Add(new_bin_no.Trim)

                        With dgv
                            item_no = .CurrentRow.Cells("ItemNo").Value
                            loc = "00" & CStr(Locations.Factory)
                            qty_on_hand = 0
                        End With

                        If sComboSource = Locations.Factory.ToString Then
                            With dtAllFactoryBins
                                Dim rw As DataRow = dtAllFactoryBins.NewRow
                                rw("item_no") = item_no
                                rw("loc") = loc
                                rw("bin_no") = new_bin_no
                                rw("qty_on_hand") = 0
                                .Rows.Add(rw)
                            End With
                        Else
                            With dtAllWarehouseBins
                                Dim rw As DataRow = dtAllFactoryBins.NewRow
                                rw("item_no") = item_no
                                rw("loc") = loc
                                rw("bin_no") = new_bin_no
                                rw("qty_on_hand") = 0
                                .Rows.Add(rw)
                            End With
                        End If

                        dgvComboBinSecond.Items.Add(new_bin_no)
                        dgvComboBinSecond.Items(0).ToString()
                    Else
                        Dim item As String = txt
                        counter = 0
                        With dgvComboBinSecond
                            For Each itm As String In .Items
                                If itm.ToString.Trim = item Then
                                    counter = 1
                                End If
                            Next
                        End With
                        If counter = 0 Then
                            dgvComboBinSecond.Items.Add(item)
                        End If

                    End If
                    dgv.CurrentRow.Cells("BinSecond").Value = new_bin_no

                Case "QtyToMove"

            End Select

            If combo Is Nothing Then
                Exit Sub
            Else
                combo.Dispose()
            End If
        Catch ex As Exception

        Finally


        End Try


    End Sub

    Private Sub dgvWorkingArea_CellValidating(sender As Object, e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvWorkingArea.CellValidating
        Dim dgv As DataGridView = CType(sender, DataGridView)

        If e.ColumnIndex = 8 Then
            If dgv.Rows(e.RowIndex).Cells(7).Value < Convert.ToInt32(e.FormattedValue) Then
                MsgBox("Transfer quantity, " & e.FormattedValue & ", is too high." & vbCrLf & vbCrLf & "The quanity available to transfer is " & dgv.Rows(e.RowIndex).Cells(7).Value.ToString.Trim & ".  ", vbOKOnly, "Transfer Quantity Exceeds Available Quantity")
                dgv.CancelEdit()
            End If
        End If
    End Sub

#End Region

#Region "  REGION - Process Inventory Transaction to Save Methods  "

    Private Sub btnDone_Click(sender As System.Object, e As System.EventArgs) Handles btnDone.Click
        WorkAreaDone()
    End Sub
    Private Sub WorkAreaDone()
        Dim ord_no As String = ""
        Dim tran_type As String = ""
        Dim distinct As Boolean = True
        Dim iBinFromQtyOnHand As Integer = 0
        Dim iBinToQtyOnHand As Integer = 0
        Dim iQtyToMove As Integer = 0

        Dim A4GLIdentity As Integer = 0
        Dim sSQL As String = ""
        'Dim cols() As String = ({"ItemNo", "TranType", "OrderNo", "TargetLocationFrom", _
        '                         "TargetLocationTo", "BinFirst", "BinSecond"})
        Dim cols() As String = ({"ItemNo", "TranType", "OrderNo", "TargetLocationFrom", _
                                        "TargetLocationTo"})

        Dim dgv As DataGridView = CType(dgvWorkingArea, DataGridView)
        Dim bin As String = ""
        Dim errtxt As String = ""

        'Get the Data from the DataGridView into a DataTable so we can work with it.  
        Dim dt As DataTable = GetDataTableFromDataGridView(dgv)
        'Make sure you have all the correct BinQtyOnHand values for the BinTo Locations '
        'and if there's nothing in the table, set rw("BinToQtyOnHand") with a 0  ....
        For Each rw As DataRow In dt.Rows

            Dim dtQtyOnHand As DataTable = cBusObj.GetQtyOnHandFromBin(rw("ItemNo").ToString.Trim, rw("TargetLocationTo").ToString.Trim, rw("BinSecond").ToString.Trim, cn)
            If dtQtyOnHand.Rows.Count = 0 Then
                rw("BinToQtyOnHand") = 0
            Else
                rw("BinToQtyOnHand") = dtQtyOnHand(0)(0)
            End If
        Next

        Dim dtItems As DataTable = dt.DefaultView.ToTable(distinct, cols)
        Dim BinFromQtyOnHand As New DataColumn("BinFromQtyOnHand", GetType(Integer))
        Dim BinToQtyOnHand As New DataColumn("BinToQtyOnHand", GetType(Integer))
        Dim QtyToMove As New DataColumn("QtyToMove", GetType(Integer))

        dtItems.Columns.Add(BinFromQtyOnHand)
        dtItems.Columns.Add(BinToQtyOnHand)
        dtItems.Columns.Add(QtyToMove)

        'create a new SQLTransacitonList
        SQLTransactionList = New List(Of String)

        If uidata.TranType = TransferType.Transfer.ToString Then
            If Not ValidataDataGridViewQty(dgv, dgv.CurrentCell.RowIndex) Then
                bin = dgv.Rows(dgv.CurrentCell.RowIndex).Cells("BinFirst").Value.ToString.Trim
                errtxt = "The quantity is more than available in Bin " & bin
                'MsgBox(errtxt, MsgBoxStyle.OkOnly, "Quantity to Move Invalid")
                lblQtyError.Text = errtxt
                Panel16.Visible = True
                Exit Sub
            Else
                errtxt = ""
                bin = ""
                Panel16.Visible = False
            End If
        End If

        'Fill in the OrderNos and OnHandQty, QtyToMove, Location and TranType Details 
        '     to test for bKeepFactoryBin variable first ...
        For i As Integer = 0 To dtItems.Rows.Count - 1
            Dim itm As String = dtItems(i)(0).ToString.Trim
            Dim trntyp As String = dtItems(i)(1).ToString.Trim
            Dim locfrom As String = dtItems(i)(3).ToString.Trim
            Dim locto As String = dtItems(i)(4).ToString.Trim

            ord_no = ""
            'the ord_no here is more like an identifier created by MACOLA for the Inventory Transaction, it's an incremented value 
            ord_no = cBusObj.ExecuteSQLScalar("select LTRIM(next_doc_no) from IMCTLFIL_SQL where im_ctl_key_1 = 1", cn).ToString.Trim
            sSQL = "Update imctlfil_sql SET  next_doc_no = (" & ord_no & " + 1)"
            cBusObj.Execute_NonSQL(sSQL, cn)
            ord_no = "00" & (Convert.ToInt32(ord_no) + 1).ToString
            dtItems(i)(2) = ord_no

            'Add Quantities From and Move
            Dim queryfrom = _
               (From items In dt.AsEnumerable().Distinct
                Where (items.Field(Of String)("ItemNo") = itm _
                       And items.Field(Of String)("TranType") = trntyp _
                       And items.Field(Of String)("TargetLocationFrom") = locfrom) _
                Group items By mItemNo = items.Field(Of String)("ItemNo"), _
                               mTranType = items.Field(Of String)("TranType"), _
                               mLocFrom = items.Field(Of String)("TargetLocationFrom"), _
                               mBinFromQtyOnHand = items.Field(Of String)("BinFromQtyOnHand"), _
                               mBinFirst = items.Field(Of String)("BinFirst") Into g = Group
            Select New With _
                   { _
                    .ItemNo = mItemNo, _
                    .TranType = mTranType, _
                    .LocFrom = mLocFrom, _
                    .BinFromQtyOnHand = mBinFromQtyOnHand
                   })

            For Each o In queryfrom.Distinct
                If o.BinFromQtyOnHand > "" Then iBinFromQtyOnHand += o.BinFromQtyOnHand
            Next

            Dim querymove = _
                            From items In dt.AsEnumerable()
                            Where (items.Field(Of String)("ItemNo") = itm _
                                   And items.Field(Of String)("TranType") = trntyp) _
                            Group items By mItemNo = items.Field(Of String)("ItemNo"), _
                                           mTranType = items.Field(Of String)("TranType") Into g = Group
                            Select New With _
                               { _
                                .QtyToMove = g.Sum(Function(items) items.Field(Of String)("QtyToMove"))
                               }
            For Each o In querymove.Distinct
                iQtyToMove += o.QtyToMove
            Next

            'Add these to dtItems Table for testing later ...
            dtItems(i)("BinFromQtyOnHand") = iBinFromQtyOnHand
            dtItems(i)("QtyToMove") = iQtyToMove
            dtItems(i)("TargetLocationFrom") = locfrom
            dtItems(i)("TargetLocationTo") = locto

        Next

        'For Each is used here to split the 'M' Item_no into groups.  
        '  dt is the DataTable created from what is in the dgvWorkingArea
        '  dtItems is the DataTable created from dt.DefaultView of the DISTINCE ItemNo, OrderNo and TranType
        '  We then populate an Array of DataRows with each ItemNo using dt.Select (ItemNo = ....)
        '  which will give us the ItemNo/OrderNo/TranType once, but the dt from the Grid will return all individual transactions.  
        '  This way the user can process more than one item/orderno/transtype combination at a time
        '  The lines of the drProcess() Array for each 'M' item_no group are split later into individual lines with drProcess 
        '  which then is used to create dtProcess datatable ...
        For Each rw As DataRow In dtItems.Rows
            ord_no = rw("OrderNo")
            tran_type = rw("TranType")
            Dim drProcess() As DataRow = dt.Select("ItemNo = '" & rw("ItemNo").ToString.Trim & "'")

            'Here is where we set bKeepFactoryBin.  We know the: 
            '     TransactionType, must be a Transfer to consider deleting
            '     Loc From - Must be from 001 (Factory) to consider deleting
            '     Loc From and To, must be the from 001 to 002 to consider deleting
            '     BinFromQtyOnHand and QtyToMove, subtracted from each other must = 0 to consider deleting

            bKeepFactoryBin = False
            If rw("TargetLocationFrom") = "002" Then '002 bins NEVER deleted and we are not emptying Factory Bin, so no need to delete anything
                bKeepFactoryBin = True
            ElseIf rw("TranType") <> "Transfer" Then 'Do do a delete if it's not a transfer.  Leave that to be handled manually if they want it
                bKeepFactoryBin = True
            ElseIf rw("TranType") = "Transfer" And rw("TargetLocationFrom") = "001" Then 'various possibilities if a Transfer and LocFrom is 001 (emptying Factory Bins)
                If rw("TargetLocationTo") = rw("TargetLocationFrom") AndAlso (drProcess(0)("BinFirst") = "FACTORY" OrElse drProcess(0)("BinFirst") = "DOCK") Then 'moving from 001 FACTORY or DOCK to another 001 bin, keep! FACTORY AND DOCK, though empty are never deleted
                    bKeepFactoryBin = True
                ElseIf Not (rw("BinFromQtyOnHand") - rw("QtyToMove") = 0) Then 'If moving from Factory to Warehouse but Factory Bin NOT Zero'd out, Keep the Factory Bin
                    bKeepFactoryBin = True
                End If
            End If

            If Not drProcess Is Nothing Then
                Dim dtProcess As DataTable = drProcess.CopyToDataTable
                If Not ProcessInventory(dtProcess, ord_no, tran_type) Then
                    SQLTransactionList = Nothing
                    Exit Sub
                End If

            End If
        Next

        bSkipPaint = True
        With Timer4
            .Interval = 250
            .Enabled = True
        End With

    End Sub

    Private Function GetDataTableFromDataGridView(dgv As DataGridView) As DataTable
        Dim dt As New DataTable()

        ' add the columns to the datatable            
        If dgv IsNot Nothing Then
            For i As Integer = 0 To dgv.Columns.Count - 1
                dt.Columns.Add(dgv.Columns(i).Name.ToString)
            Next
            dt.Columns.Add("OrderNo")
        End If

        '  add each of the data rows to the table
        For Each row As DataGridViewRow In dgv.Rows
            Dim dr As DataRow
            dr = dt.NewRow()

            For i As Integer = 0 To row.Cells.Count - 1
                dr(i) = row.Cells(i).Value.ToString.Replace(" ", "")
            Next
            dt.Rows.Add(dr)
        Next
        'Add a blank column for bKeepBinTo value used to determine if the bin goes after the transaction is complete
        Dim colKeepFactoryBin As New DataColumn("KeepFactoryBin")
        dt.Columns.Add(colKeepFactoryBin)

        Return dt
    End Function

    Private Function GetDataTableFromDataRowsCollection(rws As DataGridViewSelectedRowCollection) As DataTable
        Dim dt As New DataTable()

        Dim cls As DataGridViewColumnCollection = rws(0).DataGridView.Columns

        If cls IsNot Nothing Then
            For Each cl As DataGridViewColumn In cls
                dt.Columns.Add(cl.Name.ToString)
            Next
        End If

        For Each rw As DataGridViewRow In rws
            Dim dr As DataRow = dt.NewRow
            For i As Integer = 0 To rw.Cells.Count - 1
                dr(i) = rw.Cells(i).Value.ToString.Trim
            Next
            dt.Rows.Add(dr)
        Next

        Return dt
    End Function

    Private Function ProcessInventory(dtProcess As DataTable, orderno As String, tran_type As String) As Boolean
        'This function performs the INSERT, UPDATE to the various inventory tables.
        Dim SQLTran As String = ""
        Dim SQLCommandType As String = ""
        Dim sSQL As String
        Dim o As Object = Nothing
        Dim result As Boolean = False

        TransEffectiveDate = cBusObj.GetMacolaDate(dtpTransactionDate.Value)

        uidata = New UICurrentData

        'empty the work table ...
        sSQL = "Truncate Table IMINVBIN_WRK"
        cBusObj.Execute_NonSQL(sSQL, cn)

        'Collect the data entered by the user ....
        With dtProcess
            Dim rwnumber As Integer = 1
            For Each rw As DataRow In .Rows
                'This is the last setting for bKeepFactoryBin.  If the bin is FACTORY or DOCK, then keep it in all cases.  
                If rw("BinFirst") = "FACTORY" OrElse rw("BinFirst") = "DOCK" OrElse rw("BinFirst") = "DOCK" Then bKeepFactoryBin = True

                uidata.TranType = rw("TranType").ToString.Trim
                uidata.ItemNo = rw("ItemNo").ToString.Trim
                uidata.TargetLocFrom = rw("TargetLocationFrom").ToString.Trim
                uidata.TargetLocTo = rw("TargetLocationTo").ToString.Trim
                uidata.BinFrom = rw("BinFirst").ToString.Trim
                uidata.BinTo = rw("BinSecond").ToString.Trim
                If rw("BinFromQtyOnHand").ToString.Trim = "" Then uidata.BinFromQtyOnHand = 0 Else uidata.BinFromQtyOnHand = Convert.ToInt32(rw("BinFromQtyOnHand"))
                If rw("BinFromTransQty").ToString.Trim = "" Then uidata.BinFromTransQty = 0 Else uidata.BinFromTransQty = Convert.ToInt32(rw("BinFromTransQty"))
                If rw("BinToQtyOnHand").ToString.Trim = "" Then uidata.BinToQtyOnHand = 0 Else uidata.BinToQtyOnHand = Convert.ToInt32(rw("BinToQtyOnHand"))
                If rw("BinToTransQty").ToString.Trim = "" Then uidata.BinToTransQty = 0 Else uidata.BinToTransQty = Convert.ToInt32(rw("BinToTransQty"))
                If rw("QtyToMove").ToString.Trim = "" Then uidata.QtyToMove = 0 Else uidata.QtyToMove = Convert.ToInt32(rw("QtyToMove"))
                'UpdateIMINV tells us if we should update IMINVLOC and how to calculate IMIMVTRX.  If transaction is a Transfer and from/to are the same location (moving in the same warehouse) 
                'then no quantity change at that location.  Bin quanties change, but IMINV is Location Level and at the Location Level, the quantity is unchanged, no update. 
                If rw("TargetLocationFrom").ToString.Trim = rw("TargetLocationTo").ToString.Trim And rw("TranType").ToString.Trim = TransferType.Transfer.ToString Then
                    uidata.UpdateIMINV = "N"
                Else
                    uidata.UpdateIMINV = "Y"
                End If

                'Validate that enough data was entered, if not, do not allow further processing...
                With uidata
                    Select Case .TranType.ToString
                        Case "Transfer"
                            If uidata.BinTo.ToString.Trim = "" Then
                                MsgBox("Row # " & rwnumber.ToString & ": Factory Bin Number has not been entered.  A bin must be entered to proceed.", MsgBoxStyle.OkOnly, "Missing Factory Bin")
                                uidata.Clear()
                                Return False
                            ElseIf uidata.BinFrom.ToString.Trim = "" Then
                                MsgBox("Row # " & rwnumber.ToString & ": Warehouse Bin Number has not been entered.  A bin must be entered to proceed.", MsgBoxStyle.OkOnly, "Missing Warehouse Bin")
                                uidata.Clear()
                                Return False
                            ElseIf uidata.QtyToMove = 0 Then
                                MsgBox("Row # " & rwnumber.ToString & ": Quantity to move has not been entered.  A quantity must be entered to proceed.", MsgBoxStyle.OkOnly, "Missing Quantity")
                                uidata.Clear()
                                Return False
                            ElseIf uidata.BinFromQtyOnHand = 0 Then
                                MsgBox("Row # " & rwnumber.ToString & ": The Bin you are transferring from has no quantity.  A transfer must have a quantity in the Bin From location.", MsgBoxStyle.OkOnly, "Missing Transfer Quantity")
                                uidata.Clear()
                                Return False
                            End If
                        Case "Issue"
                            If uidata.BinFrom.ToString.Trim = "" Then
                                MsgBox("Row # " & rwnumber.ToString & ": A Bin Number has not been entered.  A bin must be entered to proceed.", MsgBoxStyle.OkOnly, "Missing Warehouse Bin")
                                uidata.Clear()
                                Return False
                            ElseIf uidata.QtyToMove = 0 Then
                                MsgBox("Row # " & rwnumber.ToString & ": Quantity to Issue has not been entered.  A quantity must be entered to proceed.", MsgBoxStyle.OkOnly, "Missing Quantity")
                                uidata.Clear()
                                Return False
                            End If
                        Case "Receipt"
                            If uidata.BinTo.ToString.Trim = "" Then
                                MsgBox("Row # " & rwnumber.ToString & ": A Bin Number has not been entered.  A bin must be entered to proceed.", MsgBoxStyle.OkOnly, "Missing Factory Bin")
                                uidata.Clear()
                                Return False
                            ElseIf uidata.QtyToMove = 0 Then
                                MsgBox("Row # " & rwnumber.ToString & ": Quantity to receive has not been entered.  A quantity must be entered to proceed.", MsgBoxStyle.OkOnly, "Missing Quantity")
                                uidata.Clear()
                                Return False
                            End If
                    End Select
                End With

                sSQL = BuildSQLString(uidata)
                cBusObj.Execute_NonSQL(sSQL, cn)
                uidata.Clear()
                rwnumber += 1
            Next
        End With

        Dim dtTransactionList As New DataTable
        dtTransactionList = cBusObj.FillIMInvBinWrk_DataTable(orderno, tran_type, cn)

        'error is here ...

        If Not Prepare_IMINVBIN_Transactions(dtTransactionList) Then Return False
        If Not Prepare_IMBINTRX_Transactions(dtTransactionList, orderno, uidata.TranType.ToString.Trim) Then Return False
        If Not Prepare_IMINVLOC_Transactions(dtTransactionList) Then Return False
        If Not Prepare_IMINVTRX_Transactions(dtTransactionList, orderno) Then Return False

        'Delete the Bin if it's from 001 AND is the FROM bin (lev_no = 0) and is a Transfer AND is Empty
        If bKeepFactoryBin = False Then
            Prepare_IMINVBIN_Delete(dtTransactionList)
            bKeepFactoryBin = True 'Delete Factory Bin is the last of the Process Transactions, so if it's False, let's set it back to Default (True)
        End If


        result = SaveTransaction(orderno)
        If result = True Then
            sSQL = "Select tran_type, item_no, loc_from, loc_to, loc_from_qty_on_hand, loc_to_qty_on_hand, bin_from, bin_to, bin_from_qty_on_hand, " & _
                   "bin_from_trx_qty, bin_to_qty_on_hand, bin_to_trx_qty, qty_to_move from IMINVBIN_WRK"
            Dim dt As DataTable = cBusObj.ExecuteSQLDataTable(sSQL, "IMINVBIN_WRK", cn)

            For Each rw As DataRow In dt.Rows
                cBusObj.SaveRepetitiveLog(rw("tran_type").ToString.Trim, rw("item_no").ToString.Trim, rw("loc_from").ToString.Trim, rw("loc_to").ToString.Trim, _
                                          rw("loc_from_qty_on_hand"), rw("loc_to_qty_on_hand"), rw("bin_from").ToString.Trim, rw("bin_to").ToString.Trim, _
                                          rw("bin_from_qty_on_hand"), rw("bin_to_qty_on_hand"), rw("bin_from_trx_qty"), rw("bin_to_trx_qty"), _
                                          rw("qty_to_move"), sessiondata.CreateDate, orderno, sessiondata.TransactionID, sessiondata.SessionID, cn)
                With sessiondata
                    .TransactionID += 1
                End With
            Next

            ''Show transaction(s) on the Transaction Log DGV
            FillTransactionLog(sessiondata.SessionID, RepetitiveLogType.SessionID)
        End If
        Return result


    End Function

    Private Sub FillTransactionLog(value As Integer, type As Integer)
        Dim sSQL As String
        Dim sWhere As String = ""

        Select Case type
            Case RepetitiveLogType.SessionID
                sWhere = "where session_id = " & value
            Case RepetitiveLogType.CreateDate
                sWhere = "where create_dt = " & value & " "
        End Select

        sSQL = "Select tran_type, item_no, loc_from, bin_from, loc_to, bin_to, bin_from_qty_on_hand, " & _
                  "bin_from_trx_qty, bin_to_qty_on_hand, bin_to_trx_qty, qty_to_move, session_id, tran_id, create_dt from IMREPLOG_MAS " & _
                  sWhere & _
                  "order by session_id, tran_id, bin_from, bin_to "

        Dim dtLog As New DataTable
        dtLog = cBusObj.ExecuteSQLDataTable(sSQL, "RepetitiveLog", cn)
        dgvTransactionLog.DataSource = dtLog

    End Sub

    Private Function Prepare_IMINVBIN_Transactions(dtTransactionList As DataTable) As Boolean
        Dim SQLTran As String = ""
        Dim SQLCommandType As String = ""
        Dim sSQL As String = ""
        Dim o As Object = Nothing

        Dim queryfrom = From trandata In dtTransactionList.AsEnumerable() _
                         Where (trandata.Field(Of Integer)("lev_no") = 0) _
                         Group trandata By TranType = trandata.Field(Of String)("tran_type"), Loc = trandata.Field(Of String)("loc"), _
                                           ItemNo = trandata.Field(Of String)("item_no"), LevelNo = trandata.Field(Of Integer)("lev_no"), _
                                           BinNo = trandata.Field(Of String)("bin_no"), BinQtyOnHand = trandata.Field(Of Integer)("bin_qty_on_hand"), _
                                           CalculationOperator = trandata.Field(Of String)("operator") Into g = Group
                             Select New With _
                                    { _
                                     .tran_type = TranType, _
                                     .item_no = ItemNo, _
                                     .loc = Loc, _
                                     .bin_no = BinNo, _
                                     .bin_qty_on_hand = BinQtyOnHand, _
                                     .total_qty_to_move = g.Sum(Function(trandata) trandata.Field(Of Integer)("qty_to_move")), _
                                     .lev_no = LevelNo, _
                                     .operator = CalculationOperator
                                     }

        Dim queryto = From trandata In dtTransactionList.AsEnumerable() _
                        Where (trandata.Field(Of Integer)("lev_no") = 1) _
                            Group trandata By TranType = trandata.Field(Of String)("tran_type"), Loc = trandata.Field(Of String)("loc"), _
                                          ItemNo = trandata.Field(Of String)("item_no"), LevelNo = trandata.Field(Of Integer)("lev_no"), _
                                          BinNo = trandata.Field(Of String)("bin_no"), BinQtyOnHand = trandata.Field(Of Integer)("bin_qty_on_hand"), _
                                          CalculationOperator = trandata.Field(Of String)("operator") Into g = Group
                            Select New With _
                                   { _
                                    .tran_type = TranType, _
                                    .item_no = ItemNo, _
                                    .loc = Loc, _
                                    .bin_no = BinNo, _
                                    .bin_qty_on_hand = BinQtyOnHand,
                                    .total_qty_to_move = g.Sum(Function(trandata) trandata.Field(Of Integer)("qty_to_move")), _
                                    .lev_no = LevelNo, _
                                    .operator = CalculationOperator
                                    }


        Dim transactions As List(Of TransactionData) = New List(Of TransactionData)

        For Each o In queryfrom

            Dim tran As New TransactionData
            With tran
                If o.tran_type = "Transfer" Then
                    .TranType = "T"
                ElseIf o.tran_type = "Issue" Then
                    .TranType = "I"
                ElseIf o.tran_type = "Receipt" Then
                    .TranType = "R"
                End If
                .LevelNo = o.lev_no
                .ItemNo = o.item_no.ToString.Trim
                .Loc = o.loc.ToString.Trim
                .BinNo = o.bin_no.ToString.Trim
                .QtyOnHand = o.bin_qty_on_hand
                .QtyToMove = o.total_qty_to_move
                .CalculationOperator = o.operator
                'Validate Quantity before proceeding; Negative inventory Qty not allowed ...
                If (.TranType = "T" Or .TranType = "I") And .QtyOnHand < .QtyToMove Then
                    MsgBox("Reduce the quantity you are saving, " & .QtyToMove & " to a value less than or equal to the quantity on hand, which is " & .QtyOnHand & ".  Then try saving again." & vbCrLf & vbCrLf & _
                           "Error writing to IMINVBIN_SQL table.  Quantity on hand in bin " & .BinNo & " at location " & .Loc & " is " & .QtyOnHand & "." & vbCrLf & vbCrLf & _
                           "The quantity being moved is " & .QtyToMove & ".  This would create a negative inventory amount which is not allowed." _
                           , MsgBoxStyle.OkOnly, "Save to IMINVBIN_SQL Error")
                    Return False
                End If

                .TransEffectiveDate = TransEffectiveDate

            End With
            transactions.Add(tran)

        Next

        'this is called only if there was a lev_no = 1, which is only in transfers.  So it will be skipped on Receipt and Issue
        For Each o In queryto

            Dim tran = New TransactionData
            With tran
                If o.tran_type = "Transfer" Then
                    .TranType = "T"
                ElseIf o.tran_type = "Issue" Then
                    .TranType = "I"
                ElseIf o.tran_type = "Receipt" Then
                    .TranType = "R"
                End If
                .LevelNo = o.lev_no
                .ItemNo = o.item_no.ToString.Trim
                .Loc = o.loc.ToString.Trim
                .BinNo = o.bin_no.ToString.Trim
                .QtyOnHand = o.bin_qty_on_hand
                .QtyToMove = o.total_qty_to_move
                .CalculationOperator = o.operator
                .TransEffectiveDate = TransEffectiveDate
            End With
            transactions.Add(tran)

        Next

        For Each transaction As TransactionData In transactions

            Dim sBinExists As String
            sSQL = ""
            sSQL = " Select item_no from IMINVBIN_SQL where RTrim(LTrim(item_no)) = '" & transaction.ItemNo & "' and RTrim(LTrim(bin_no)) = '" & transaction.BinNo & "' and RTrim(LTrim(loc)) = '" & transaction.Loc & "'"
            sBinExists = Convert.ToString(cBusObj.ExecuteSQLScalar(sSQL, cn)).Trim
            SQLCommandType = IIf(Not (sBinExists Is Nothing OrElse sBinExists = ""), SQLCommandTypeEnum.UPDATE.ToString, SQLCommandTypeEnum.INSERT.ToString)


            With transaction
                invbin = New InventoryBin
                If SQLCommandType = SQLCommandTypeEnum.UPDATE.ToString Then
                    invbin = invbin.GetInventoryBin(.ItemNo, .BinNo, .Loc, cn)
                    .BinPriority = invbin.bin_priority
                    .BinStatus = invbin.bin_status
                Else
                    .BinPriority = 98
                    .BinStatus = "A"
                    invbin.issue_priority = 1
                End If

                'HOW TO DETERMINE .ReceivedDt in queryfrom:  Receipts (receive to a bin) and Transfers lev_no = 1 use TransEffectiveDate, everything else .received_dt
                'Received Date:  
                If .TranType = "R" Or SQLCommandType = SQLCommandTypeEnum.INSERT.ToString Or .LevelNo = "1" Then
                    .ReceivedDt = .TransEffectiveDate
                Else
                    Dim sql As String = "Select received_dt from IMINVBIN_SQL Where item_no = '" & .ItemNo & "' and bin_no = '" & .BinNo & "' and loc = '" & .Loc & "' "
                    .ReceivedDt = cBusObj.ExecuteSQLScalar(sql, cn)
                End If

            End With


            SQLTran = cBusObj.PrepareInventoryBinSQLTransaction(transaction.TranType, transaction, Bins.BinFirst.ToString, SQLCommandType, cn, invbin)


            'store the SQL String to the list so we can write all of them at once and use the BEGIN TRANSACTION / COMMIT / ROLLBACK functionality ...
            SQLTransactionList.Add(SQLTran)

            ''Just setting bDeleteBin here for the last SQLTran to delete the bin if we are coming out of location 001 (FACTORY) 
            ''and the lev_no = 0 (which means it's the FROM bin).  The Delete transaction is the last call in ProcessInventory

            'If bKeepFactoryBin = False Then 'This overrides the bDeleteBin Check, since is has already determined it is a Factory to Factory Transfer
            '    If bDeleteBin = False Then
            '        bDeleteBin = IIf(transaction.Loc = "001" And transaction.LevelNo = 0 And transaction.QtyOnHand - transaction.QtyToMove = 0 _
            '                     And transaction.TranType = "T", True, False)
            '    End If
            'End If

        Next

        Return True

    End Function

    Private Function Prepare_IMBINTRX_Transactions(dtTransactionList As DataTable, orderno As String, TransType As String) As Boolean

        '--== 2 ==-- Inventory Bin Trx Table ...

        Dim SQLTran As String = ""
        Dim SQLCommandType As String = ""
        Dim sSQL As String = ""
        Dim o As Object = Nothing

        Dim queryfrom = From trandata In dtTransactionList.AsEnumerable() _
                        Where (trandata.Field(Of Integer)("lev_no") = 0)
                        Group trandata By TranType = trandata.Field(Of String)("tran_type"), Loc = trandata.Field(Of String)("loc"), _
                                          ItemNo = trandata.Field(Of String)("item_no"), LevelNo = trandata.Field(Of Integer)("lev_no"), _
                                          BinNo = trandata.Field(Of String)("bin_no"), UserName = trandata.Field(Of String)("user_name") _
                                          Into g = Group
                            Select New With _
                                   { _
                                    .tran_type = TranType, _
                                    .item_no = ItemNo, _
                                    .loc = Loc, _
                                    .bin_no = BinNo, _
                                    .bin_qty_on_hand = g.Sum(Function(trandata) trandata.Field(Of Integer)("bin_qty_on_hand")), _
                                    .total_qty_to_move = g.Sum(Function(trandata) trandata.Field(Of Integer)("qty_to_move")), _
                                    .lev_no = LevelNo, _
                                    .user_name = UserName
                                    }

        Dim queryto = From trandata In dtTransactionList.AsEnumerable() _
                        Where (trandata.Field(Of Integer)("lev_no") = 1) _
                        Group trandata By TranType = trandata.Field(Of String)("tran_type"), Loc = trandata.Field(Of String)("loc"), _
                                          ItemNo = trandata.Field(Of String)("item_no"), LevelNo = trandata.Field(Of Integer)("lev_no"), _
                                          BinNo = trandata.Field(Of String)("bin_no"), UserName = trandata.Field(Of String)("user_name") _
                                          Into g = Group
                            Select New With _
                                   { _
                                    .tran_type = TranType, _
                                    .item_no = ItemNo, _
                                    .loc = Loc, _
                                    .bin_no = BinNo, _
                                    .bin_qty_on_hand = g.Sum(Function(trandata) trandata.Field(Of Integer)("bin_qty_on_hand")),
                                    .total_qty_to_move = g.Sum(Function(trandata) trandata.Field(Of Integer)("qty_to_move")), _
                                    .lev_no = LevelNo, _
                                    .user_name = UserName
                                    }

        Dim transactions As List(Of TransactionData) = New List(Of TransactionData)
        For Each o In queryfrom
            Dim tran As New TransactionData
            With tran
                .TranType = o.tran_type.ToString.Trim
                .LevelNo = o.lev_no
                .ItemNo = o.item_no.ToString.Trim
                .Loc = o.loc.ToString.Trim
                .BinNo = o.bin_no.ToString.Trim
                .QtyOnHand = o.bin_qty_on_hand 'NOTE: QtyOnHand is retrieved only to validate the QtyToMove is not a negative number, it's not written to the IMBINTRX table
                .QtyToMove = o.total_qty_to_move
                .OrderNo = orderno.ToString.Trim
                .UserName = o.user_name

                'Validate Quantity before proceeding; Negative inventory Qty not allowed ...
                If Not (.TranType = "Receipt") AndAlso .QtyOnHand < .QtyToMove Then
                    MsgBox("Error writing to IMINVBIN_SQL table.  Quantity on hand in bin " & .BinNo & "at location " & .Loc & " is " & .QtyOnHand & "." & vbCrLf & vbCrLf & _
                           "The quantity being moved is " & .QtyToMove & ".  This would create a negative inventory amount which is not allowed.", MsgBoxStyle.OkOnly, "Save to IMINVBIN_SQL Error")
                    Return False
                End If
                .TransEffectiveDate = TransEffectiveDate
            End With
            transactions.Add(tran)
        Next
        For Each o In queryto
            Dim tran = New TransactionData
            With tran
                .TranType = o.tran_type.ToString.Trim
                .LevelNo = o.lev_no
                .ItemNo = o.item_no.ToString.Trim
                .Loc = o.loc.ToString.Trim
                .BinNo = o.bin_no.ToString.Trim
                .QtyOnHand = o.bin_qty_on_hand
                .QtyToMove = o.total_qty_to_move
                .OrderNo = orderno.ToString.Trim
                .UserName = o.user_name
                .TransEffectiveDate = TransEffectiveDate
                transactions.Add(tran)
            End With
        Next

        For Each transaction As TransactionData In transactions
            'Only INSERT for IMBINTRX_SQL table...
            SQLTran = cBusObj.PrepareBinTrxSQL(transaction, cn)
            SQLTransactionList.Add(SQLTran)
        Next

        Return True
    End Function

    Private Function Prepare_IMINVLOC_Transactions(dtTransactionList As DataTable) As Boolean
        Dim SQLTran As String = ""
        Dim SQLCommandType As String = ""
        Dim sSQL As String = ""

        '--== 3 ==-- Inventory Location Table ...
        If dtTransactionList.Rows(0)("update_iminv").ToString = "N" Then Return True
        Dim queryfrom = _
            From trandata In dtTransactionList.AsEnumerable()
            Where (trandata.Field(Of Integer)("lev_no") = 0) _
            Group trandata By TranType = trandata.Field(Of String)("tran_type"), _
                              Loc = trandata.Field(Of String)("loc"), _
                              LocQtyOnHand = trandata.Field(Of Integer)("loc_qty_on_hand"), _
                              ItemNo = trandata.Field(Of String)("item_no"), _
                              LevelNo = trandata.Field(Of Integer)("lev_no"), _
                              CalculationOperator = trandata.Field(Of String)("operator") Into g = Group
            Select New With _
                { _
                .trantype = TranType, _
                .loc = Loc, _
                .loc_qty_on_hand = LocQtyOnHand, _
                .itemno = ItemNo, _
                .bin_qty_on_hand = g.Sum(Function(trandata) trandata.Field(Of Integer)("bin_qty_on_hand")), _
                .total_qty_to_move = g.Sum(Function(trandata) trandata.Field(Of Integer)("qty_to_move")), _
                .levelno = LevelNo, _
                .operator = CalculationOperator
                }

        Dim queryto = _
            From trandata In dtTransactionList.AsEnumerable()
            Where (trandata.Field(Of Integer)("lev_no") = 1) _
            Group trandata By TranType = trandata.Field(Of String)("tran_type"), _
                              Loc = trandata.Field(Of String)("loc"), _
                              LocQtyOnHand = trandata.Field(Of Integer)("loc_qty_on_hand"), _
                              ItemNo = trandata.Field(Of String)("item_no"), _
                              LevelNo = trandata.Field(Of Integer)("lev_no"), _
                              CalculationOperator = trandata.Field(Of String)("operator") Into g = Group
            Select New With _
                { _
                .trantype = TranType, _
                .loc = Loc, _
                .loc_qty_on_hand = LocQtyOnHand, _
                .itemno = ItemNo, _
                .bin_qty_on_hand = g.Sum(Function(trandata) trandata.Field(Of Integer)("bin_qty_on_hand")), _
                .total_qty_to_move = g.Sum(Function(trandata) trandata.Field(Of Integer)("qty_to_move")), _
                .levelno = LevelNo, _
                .operator = CalculationOperator
                }

        Dim transactions As List(Of TransactionData) = New List(Of TransactionData)


        For Each o In queryfrom
            Dim tran As New TransactionData
            With tran
                .TranType = o.trantype
                .LevelNo = (o.levelno)
                .ItemNo = o.itemno
                .Loc = o.loc
                .QtyOnHand = o.loc_qty_on_hand
                .QtyToMove = o.total_qty_to_move
                'Validate Quantity before proceeding; Negative inventory Qty not allowed ...
                If Not (.TranType = "Receipt") AndAlso .QtyOnHand < .QtyToMove Then
                    MsgBox("Error writing to IMINVBIN_SQL table.  Quantity on hand in bin " & .BinNo & "at location " & .Loc & " is " & .QtyOnHand & "." & vbCrLf & vbCrLf & _
                           "The quantity being moved is " & .QtyToMove & ".  This would create a negative inventory amount which is not allowed.", MsgBoxStyle.OkOnly, "Save to IMINVBIN_SQL Error")
                    Return False
                End If
                .CalculationOperator = o.operator
            End With
            transactions.Add(tran)
        Next


        For Each o In queryto
            Dim tran = New TransactionData
            With tran
                .TranType = o.trantype
                .LevelNo = (o.levelno)
                .ItemNo = o.itemno
                .Loc = o.loc
                .QtyOnHand = o.loc_qty_on_hand
                .QtyToMove = o.total_qty_to_move
                .CalculationOperator = o.operator
                transactions.Add(tran)
            End With

        Next

        For Each transaction As TransactionData In transactions

            'This Select is to get the full IMINVLOC_SQL record and set the Usage PTD and YTD values (see below)
            Select Case transaction.TranType
                Case "Transfer", "Receipt"
                    invloc = New InventoryLocation
                    invloc = invloc.GetInventoryLocation(transaction.ItemNo, transaction.Loc, cn)
                    transaction.UsagePTD = invloc.usage_ptd ' If Transfer, then Usuage stays the same for PTD and YTD.  
                    transaction.UsageYTD = invloc.usage_ytd
                Case "Issue"
                    invloc = New InventoryLocation
                    invloc = invloc.GetInventoryLocation(transaction.ItemNo, transaction.Loc, cn)
                    transaction.UsagePTD = (invloc.usage_ptd + transaction.QtyToMove) ' If it's Issued, the Usage has occurred, so we add to the PTD and YTD
                    transaction.UsageYTD = (invloc.usage_ytd + transaction.QtyToMove)
            End Select

            SQLTran = cBusObj.PrepareInvLocSQL(transaction, SQLCommandTypeEnum.UPDATE.ToString, cn, invloc)
            SQLTransactionList.Add(SQLTran)

        Next

        Return True
    End Function

    Private Function Prepare_IMINVTRX_Transactions(dtTransactionList As DataTable, orderno As String) As Boolean
        Dim SQLTran As String = ""
        Dim SQLCommandType As String = ""
        Dim sSQL As String = ""
        '--== 4 ==-- Inventory Transaction Table ...

        Dim queryfrom = _
        From trandata In dtTransactionList.AsEnumerable()
        Where (trandata.Field(Of Integer)("lev_no") = 0) _
        Group trandata By TranType = trandata.Field(Of String)("tran_type"), Loc = trandata.Field(Of String)("loc"), _
                                     ItemNo = trandata.Field(Of String)("item_no"), LevelNo = trandata.Field(Of Integer)("lev_no"), _
                                     ReceivedDt = trandata.Field(Of Integer)("received_dt"), IDNo = trandata.Field(Of String)("id_no"), _
                                     UserName = trandata.Field(Of String)("user_name"), _
                                     CalculationOperator = trandata.Field(Of String)("operator"), _
                                     UpdateIMINV = trandata.Field(Of String)("update_iminv") Into g = Group
        Select New With _
            { _
            .trantype = TranType, _
            .loc = Loc, _
            .itemno = ItemNo, _
            .bin_qty_on_hand = g.Sum(Function(trandata) trandata.Field(Of Integer)("bin_qty_on_hand")), _
            .total_qty_to_move = g.Sum(Function(trandata) trandata.Field(Of Integer)("qty_to_move")), _
            .levelno = LevelNo, _
            .user_name = UserName, _
            .id_no = IDNo,
            .operator = CalculationOperator,
            .update_iminv = UpdateIMINV
            }

        Dim queryto = _
            From trandata In dtTransactionList.AsEnumerable()
            Where (trandata.Field(Of Integer)("lev_no") = 1) _
            Group trandata By TranType = trandata.Field(Of String)("tran_type"), Loc = trandata.Field(Of String)("loc"), _
                                     ItemNo = trandata.Field(Of String)("item_no"), LevelNo = trandata.Field(Of Integer)("lev_no"), _
                                     ReceivedDt = trandata.Field(Of Integer)("received_dt"), IDNo = trandata.Field(Of String)("id_no"), _
                                     UserName = trandata.Field(Of String)("user_name"), _
                                     CalculationOperator = trandata.Field(Of String)("operator"), _
                                     UpdateIMINV = trandata.Field(Of String)("update_iminv") Into g = Group
            Select New With _
                { _
                .trantype = TranType, _
                .loc = Loc, _
                .itemno = ItemNo, _
                .bin_qty_on_hand = g.Sum(Function(trandata) trandata.Field(Of Integer)("bin_qty_on_hand")), _
                .total_qty_to_move = g.Sum(Function(trandata) trandata.Field(Of Integer)("qty_to_move")), _
                .levelno = LevelNo, _
                .user_name = UserName, _
                .id_no = IDNo,
                .operator = CalculationOperator,
                .update_iminv = UpdateIMINV
                }

        Dim transactions As List(Of TransactionData) = New List(Of TransactionData)

        For Each o In queryfrom

            Dim tran As New TransactionData
            With tran
                .TranType = o.trantype
                .LevelNo = (o.levelno)
                .ItemNo = o.itemno
                .Loc = o.loc
                .QtyOnHand = o.bin_qty_on_hand
                .QtyToMove = o.total_qty_to_move
                .OrderNo = orderno
                .UserName = o.user_name
                .IDNo = o.id_no
                .TransEffectiveDate = TransEffectiveDate
                .CalculationOperator = o.operator
                .UpdateIMINV = o.update_iminv
            End With
            transactions.Add(tran)
        Next

        For Each o In queryto

            Dim tran = New TransactionData
            With tran
                .TranType = o.trantype
                .LevelNo = (o.levelno)
                .ItemNo = o.itemno
                .Loc = o.loc
                .QtyOnHand = o.bin_qty_on_hand
                .QtyToMove = o.total_qty_to_move
                .OrderNo = orderno
                .UserName = o.user_name
                .IDNo = o.id_no
                .TransEffectiveDate = TransEffectiveDate
                .CalculationOperator = o.operator
                .UpdateIMINV = o.update_iminv
            End With
            transactions.Add(tran)

        Next

        For Each transaction As TransactionData In transactions
            invtrx = New InventoryTrx
            invtrx = invtrx.GetInventoryLocation(transaction.ItemNo, transaction.Loc, cn)

            SQLTran = cBusObj.PrepareInvTrxSQL(transaction, invtrx, cn)
            SQLTransactionList.Add(SQLTran)

        Next

        Return True

    End Function

    Private Function Prepare_IMINVBIN_Delete(dtTransactionList As DataTable) As Boolean
        Dim SQLTran As String = ""
        Dim SQLCommandType As String = ""
        Dim sSQL As String = ""
        Dim o As Object = Nothing

        Dim queryfrom = From trandata In dtTransactionList.AsEnumerable() _
                        Where (trandata.Field(Of Integer)("lev_no") = 0) _
                        Group trandata By TranType = trandata.Field(Of String)("tran_type"), Loc = trandata.Field(Of String)("loc"), _
                                          ItemNo = trandata.Field(Of String)("item_no"), LevelNo = trandata.Field(Of Integer)("lev_no"), _
                                          BinNo = trandata.Field(Of String)("bin_no"), BinQtyOnHand = trandata.Field(Of Integer)("bin_qty_on_hand"), _
                                          CalculationOperator = trandata.Field(Of String)("operator") Into g = Group
                            Select New With _
                                   { _
                                    .tran_type = TranType, _
                                    .item_no = ItemNo, _
                                    .loc = Loc, _
                                    .bin_no = BinNo, _
                                    .bin_qty_on_hand = BinQtyOnHand, _
                                    .total_qty_to_move = g.Sum(Function(trandata) trandata.Field(Of Integer)("qty_to_move")), _
                                    .lev_no = LevelNo, _
                                    .operator = CalculationOperator
                                    }


        Dim transactions As List(Of TransactionData) = New List(Of TransactionData)

        For Each o In queryfrom

            Dim tran As New TransactionData
            With tran
                If o.tran_type = "Transfer" Then
                    .TranType = "T"
                ElseIf o.tran_type = "Issue" Then
                    .TranType = "I"
                ElseIf o.tran_type = "Receipt" Then
                    .TranType = "R"
                End If
                .LevelNo = o.lev_no
                .ItemNo = o.item_no.ToString.Trim
                .Loc = o.loc.ToString.Trim
                .BinNo = o.bin_no.ToString.Trim
                .QtyOnHand = o.bin_qty_on_hand
                .QtyToMove = o.total_qty_to_move
                .CalculationOperator = o.operator
                .TransEffectiveDate = TransEffectiveDate
            End With
            transactions.Add(tran)
        Next

        For Each transaction As TransactionData In transactions

            Dim sBinExists As String
            sSQL = ""
            sSQL = " Select item_no from IMINVBIN_SQL where RTrim(LTrim(item_no)) = '" & transaction.ItemNo & "' and RTrim(LTrim(bin_no)) = '" & transaction.BinNo & "' and RTrim(LTrim(loc)) = '" & transaction.Loc & "'"
            sBinExists = Convert.ToString(cBusObj.ExecuteSQLScalar(sSQL, cn)).Trim

            If sBinExists Is Nothing OrElse sBinExists = "" Then
                MsgBox("The Factory Bin, " & transaction.BinNo & " where the Transfer quantity is coming from does not exist or another error has occured.  Contact Support with the Item No, Bin No.", _
                           MsgBoxStyle.OkOnly, "Save to IMINVBIN_SQL Error")
                Return False
            End If

            With transaction
                SQLTran = "Delete From IMINVBIN_SQL where loc = '" & .Loc.Trim & "' " & vbCrLf & _
                          "   and item_no = '" & .ItemNo.Trim & "' " & vbCrLf & _
                          "   and bin_no = '" & .BinNo.Trim & "' "
            End With

            'store the SQL String to the list so we can write all of them at once and use the BEGIN TRANSACTION / COMMIT / ROLLBACK functionality ...
            SQLTransactionList.Add(SQLTran)

        Next

        Return True

    End Function

    Private Function BuildSQLString(uidata As UICurrentData) As String
        Dim ssql As New StringBuilder

        With ssql
            .Append("INSERT INTO IMINVBIN_WRK ")
            .Append("VALUES ('" & uidata.TranType & "', ")
            .Append("'" & uidata.ItemNo & "', ")
            .Append("'" & uidata.TargetLocFrom & "', ")
            .Append("'" & uidata.TargetLocTo & "', ")
            .Append("0, ")
            .Append("0, ")
            .Append("'" & uidata.BinFrom & "', ")
            .Append("'" & uidata.BinTo & "', ")
            .Append(" " & uidata.BinFromQtyOnHand & ", ")
            .Append(" " & uidata.BinFromTransQty & ", ")
            .Append(" " & uidata.BinToQtyOnHand & ", ")
            .Append(" " & uidata.BinToTransQty & ", ")
            .Append(" " & uidata.QtyToMove & ", ")
            .Append("'" & uidata.UpdateIMINV & "')")
        End With

        Return ssql.ToString
    End Function

    Private Function SaveTransaction(ord_no As String) As Boolean
        Dim trn As SqlTransaction
        Dim cmd As SqlCommand

        trn = cn.BeginTransaction
        For Each ssql In SQLTransactionList
            cmd = New SqlCommand(ssql, cn, trn)
            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                Dim err1 As String = ("Commit Exception Type: " & ex.GetType().ToString)
                Dim err2 As String = "  Message: " & ex.Message
                MsgBox(err1 & err2, MsgBoxStyle.OkOnly, "Commit Failed")
                Try
                    trn.Rollback()
                    Return False
                    'Exit Function
                Catch ex2 As Exception
                    Dim err3 As String = ("Rollback Exception Type: " & ex.GetType().ToString)
                    Dim err4 As String = "  Message: " & ex.Message
                    MsgBox(err3 & err4, MsgBoxStyle.OkOnly, "Rollback Failed")
                    Return False
                    'Exit Function
                End Try
            End Try

        Next

        Try

            ''Do the rollback for testing here. 
            'trn.Rollback()
            'Exit Sub
            trn.Commit()
            'Dim next_doc_no = Convert.ToInt32(ord_no) + 1
            'SetIMCtlNextDocNo(next_doc_no)
        Catch ex As Exception
            Dim err1 As String = ("Commit Exception Type: " & ex.GetType().ToString)
            Dim err2 As String = "  Message: " & ex.Message
            MsgBox(err1 & err2, MsgBoxStyle.OkOnly, "Commit Failed")
            Try
                trn.Rollback()
                Return False
                'Exit Function
            Catch ex2 As Exception
                Dim err3 As String = ("Rollback Exception Type: " & ex.GetType().ToString)
                Dim err4 As String = "  Message: " & ex.Message
                MsgBox(err3 & err4, MsgBoxStyle.OkOnly, "Rollback Failed")
                Return False
                'Exit Function
            End Try

        End Try

        dgvItems.DataSource = Nothing
        dgvFactory001.DataSource = Nothing
        dgvWarehouse002.DataSource = Nothing

        ClearDataTables()
        SQLTransactionList = Nothing

        ' LoadEmptyDataTables()
        LoadItemDataGridView()
        LoadFactoryDataGridView()
        LoadWarehouseDataGridView()
        bAppLoading = True
        LoadData(bAppLoading)
        LoadData(False, uidata.ItemNo)
        dgvWorkingArea.Rows.Clear()
        Return True
    End Function

    Private Sub ClearDataTables()
        dtAllItems.Clear()
        dtAllFactoryBins.Clear()
        dtAllWarehouseBins.Clear()
        dtItem.Clear()
        dtFactory001.Clear()
        dtWarehouse002.Clear()

    End Sub

    Private Sub SetIMCtlNextDocNo(next_doc_no As Integer)
        Dim sSQL As String = "Update IMCTLFIL_SQL Set next_doc_no = " & next_doc_no & " where im_ctl_key_1 = " & IM_CTL_KEY_1.Massarelli
        cBusObj.Execute_NonSQL(sSQL, cn)
    End Sub

#End Region

    Private Sub cboDefaultFactoryBin_DropDownClosed(sender As Object, e As System.EventArgs) Handles cboDefaultFactoryBin.DropDownClosed

    End Sub

    Private Sub cboDefaultFactoryBin_MouseWheel(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles cboDefaultFactoryBin.MouseWheel
        If Not cboDefaultFactoryBin.DroppedDown Then
            Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
            mwe.Handled = True
        End If
    End Sub

    Private Sub Timer4_Tick(sender As Object, e As System.EventArgs) Handles Timer4.Tick
        Timer4.Enabled = False
        With txtItemSearch
            .Focus()
            .SelectAll()
        End With
    End Sub

    Private Sub Timer5_Tick(sender As Object, e As System.EventArgs) Handles Timer5.Tick
        Timer5.Enabled = False
        iDoneCounter = 0
    End Sub

    Private Sub Timer6_Tick(sender As System.Object, e As System.EventArgs) Handles Timer6.Tick
        Timer6.Enabled = False
        CalculateQuantity()
    End Sub
End Class

Public Class UICurrentData
    Public TranType As String
    Public ItemNo As String
    Public ItemCount As Integer
    Public TargetLocFrom As String
    Public TargetLocTo As String
    Public BinFrom As String
    Public BinTo As String
    Public BinFromQtyOnHand As Integer
    Public BinFromTransQty As Integer
    Public BinToQtyOnHand As Integer
    Public BinToTransQty As Integer
    Public QtyToMove As Integer
    Public Direction As String
    Public ComboLocationFromValue As String
    Public ComboLocationToValue As String
    Public ComboBinFirstValue As String
    Public ComboBinSecondValue As String
    Public WarehouseBinCount As Integer
    Public UpdateIMINV As String

    Public Sub Clear()
        TranType = ""
        ItemNo = ""
        ItemCount = 0
        TargetLocFrom = ""
        TargetLocTo = ""
        BinFrom = ""
        BinTo = ""
        BinFromQtyOnHand = 0
        BinFromTransQty = 0
        BinToQtyOnHand = 0
        BinToTransQty = 0
        QtyToMove = 0
        Direction = ""
        WarehouseBinCount = 0
        UpdateIMINV = "Y"
    End Sub

    Public Sub New(sTranType As String, sItemNo As String, iItemCount As Integer, sTargetLocFrom As String, sTargetLocTo As String, sBinFrom As String, _
                   sBinTo As String, iBinFromQtyOnHand As Integer, iBinFromTransQty As Integer, sDirection As String, _
                   iBinToQtyOnHand As Integer, iBinToTransQty As Integer, mWarehouseBinCount As Integer, mUpdateIMINV As String)
        TranType = sTranType
        ItemNo = sItemNo
        ItemCount = iItemCount
        TargetLocFrom = sTargetLocFrom
        TargetLocTo = sTargetLocTo
        BinFrom = sBinFrom
        BinTo = sBinTo
        BinFromQtyOnHand = 0
        BinFromTransQty = 0
        BinToQtyOnHand = 0
        BinToTransQty = 0
        QtyToMove = 0
        Direction = sDirection
        WarehouseBinCount = mWarehouseBinCount
        UpdateIMINV = mUpdateIMINV
    End Sub

    Public Sub New()
        TranType = ""
        ItemNo = ""
        ItemCount = 0
        TargetLocFrom = ""
        TargetLocTo = ""
        BinFrom = ""
        BinTo = ""
        BinFromQtyOnHand = 0
        BinFromTransQty = 0
        BinToQtyOnHand = 0
        BinToTransQty = 0
        QtyToMove = 0
        Direction = ""
        WarehouseBinCount = 0
        UpdateIMINV = "Y"
    End Sub

End Class

Public Class CurrentSessionData
    Public SessionID As Integer
    Public TransactionID As Integer
    Public CreateDate As Integer
End Class

Public Class TransactionData
    Public ItemNo As String
    Public Loc As String
    Public BinNo As String
    Public QtyOnHand As Integer
    Public QtyToMove As Integer
    Public BinPriority As Integer
    Public BinStatus As String
    Public CalculationOperator As String
    Public IDNo As String
    Public LevelNo As Integer
    Public NewBinQtyOnHand As Integer
    Public NewInvQtyOnHand As Integer
    Public OldBinQtyOnHand As Integer
    Public OldInvQtyOnHand As Integer
    Public OrderNo As String
    Public PriorYearUsage As Integer
    Public Qty As Integer
    Public TransactionQty As Integer
    Public TransEffectiveDate As Integer
    Public TranType As String
    Public UsagePTD As Decimal
    Public UsageYTD As Decimal
    Public UserName As String
    Public ReceivedDt As Integer
    Public UpdateIMINV As String

    Public Sub Clear()
        BinNo = ""
        BinPriority = 0
        BinStatus = ""
        CalculationOperator = ""
        IDNo = ""
        LevelNo = 0
        Loc = ""
        NewBinQtyOnHand = 0
        NewInvQtyOnHand = 0
        OldBinQtyOnHand = 0
        OldInvQtyOnHand = 0
        OrderNo = ""
        Qty = 0
        QtyOnHand = 0
        QtyToMove = 0
        ReceivedDt = 0
        TransactionQty = 0
        TransEffectiveDate = 0
        TranType = ""
        UsagePTD = 0
        UsageYTD = 0
        UserName = ""
        ItemNo = ""
        UpdateIMINV = "T"
    End Sub

    Public Sub New(mItemNo As String, mLoc As String, mBinNo As String, mBinPriority As Integer, mBinStatus As String, _
                   mQtyOnHand As Integer, mQtyToMove As Integer, mReceivedDt As Integer, mQty As Integer, _
                   mTransactionQty As Integer, mNewBinQtyOnHand As Integer, mOldBinQtyOnHand As Integer, _
                   mOldInvQtyOnHand As Integer, mNewInvQtyOnHand As Integer, mLevelNo As Integer, mOrderNo As String, _
                   mIDNo As String, mUserName As String, mTranType As String, mCalculationOperator As String, _
                   mTransEffectiveDate As Integer, mUsagePTD As Decimal, mUsageYTD As Decimal, mUpdateIMINV As String)

        ItemNo = mItemNo
        Loc = mLoc
        BinNo = mBinNo
        BinPriority = mBinPriority
        BinStatus = mBinStatus
        QtyOnHand = mQtyOnHand
        QtyToMove = mQtyToMove

        ReceivedDt = mReceivedDt
        Qty = mQty
        TransactionQty = mTransactionQty
        NewBinQtyOnHand = mNewBinQtyOnHand
        OldBinQtyOnHand = mOldBinQtyOnHand
        OldInvQtyOnHand = mOldInvQtyOnHand
        NewInvQtyOnHand = mNewInvQtyOnHand
        LevelNo = mLevelNo
        OrderNo = mOrderNo
        IDNo = mIDNo
        UserName = mUserName
        TranType = mTranType
        CalculationOperator = mCalculationOperator
        TransEffectiveDate = mTransEffectiveDate
        UsagePTD = mUsagePTD
        UsageYTD = mUsageYTD
        UpdateIMINV = mUpdateIMINV
    End Sub

    Public Sub New()
        ItemNo = ""
        Loc = ""
        BinNo = ""
        BinPriority = 0
        BinStatus = ""
        QtyOnHand = 0
        QtyToMove = 0
        ReceivedDt = 0
        Qty = 0
        TransactionQty = 0
        NewBinQtyOnHand = 0
        OldBinQtyOnHand = 0
        OldInvQtyOnHand = 0
        NewInvQtyOnHand = 0
        LevelNo = 0
        OrderNo = ""
        IDNo = ""
        UserName = ""
        TranType = ""
        CalculationOperator = ""
        TransEffectiveDate = 0
        UsagePTD = 0
        UsageYTD = 0
        UpdateIMINV = "T"
    End Sub

End Class

