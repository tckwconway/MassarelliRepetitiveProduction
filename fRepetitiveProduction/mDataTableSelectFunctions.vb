Module mDataTableSelectFunctions
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' This Function returns a dataTable that contains distinct values
    ''' </summary>
    ''' <param name="SourceTable">dataTable to be filtered for Distinct Selected values</param>
    ''' <param name="FieldNames">String of Column Names of the Source DataTable to be filtered and to be displayed</param>
    ''' <returns> a dataTable that contains distinct values</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    '''    [Victor]   10/25/2006   Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Friend Function SelectDistinct(ByVal SourceTable As DataTable, ByVal ParamArray FieldNames() As String) As DataTable
        Dim lastValues() As Object
        Dim newTable As DataTable

        If FieldNames Is Nothing OrElse FieldNames.Length = 0 Then
            Throw New ArgumentNullException("FieldNames")
        End If

        lastValues = New Object(FieldNames.Length - 1) {}
        newTable = New DataTable

        For Each field As String In FieldNames
            newTable.Columns.Add(field, SourceTable.Columns(field).DataType)
        Next

        For Each Row As DataRow In SourceTable.Select("", String.Join(", ", FieldNames))
            If Not fieldValuesAreEqual(lastValues, Row, FieldNames) Then
                newTable.Rows.Add(createRowClone(Row, newTable.NewRow(), FieldNames))

                setLastValues(lastValues, Row, FieldNames)
            End If
        Next

        Return newTable
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' This function returns a boolean value to test whether 2 values are equal or not
    ''' </summary>
    ''' <param name="lastValues">The last value to being tested</param>
    ''' <param name="currentRow">The current dataRow being tested</param>
    ''' <param name="fieldNames">The fieldNames in the DataRow</param>
    ''' <returns>A boolean value to test whether 2 values are equal or not</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    '''    [Victor]   10/25/2006   Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function fieldValuesAreEqual(ByVal lastValues() As Object, ByVal currentRow As DataRow, ByVal fieldNames() As String) As Boolean
        Dim areEqual As Boolean = True

        For i As Integer = 0 To fieldNames.Length - 1
            If lastValues(i) Is Nothing OrElse Not lastValues(i).Equals(currentRow(fieldNames(i))) Then
                areEqual = False
                Exit For
            End If
        Next

        Return areEqual
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' This function returns a datarow created from the source dataRow to be populated with distinct values
    ''' </summary>
    ''' <param name="sourceRow">Source DataRow</param>
    ''' <param name="newRow">Target DataRow</param>
    ''' <param name="fieldNames">Column Names in the DataRow</param>
    ''' <returns>A datarow created from the source dataRow to be populated with distinct values</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    '''    [Victor]   10/25/2006   Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function createRowClone(ByVal sourceRow As DataRow, ByVal newRow As DataRow, ByVal fieldNames() As String) As DataRow
        For Each field As String In fieldNames
            newRow(field) = sourceRow(field)
        Next

        Return newRow
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' This is a method that is used to set the last values in a DataRow
    ''' </summary>
    ''' <param name="lastValues">An object array that contains that will contain last values</param>
    ''' <param name="sourceRow">source datarow that contains the last values to be imported</param>
    ''' <param name="fieldNames">Column Names of the Source Row</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    '''    [Victor]   10/25/2006   Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub setLastValues(ByVal lastValues() As Object, ByVal sourceRow As DataRow, ByVal fieldNames() As String)
        For i As Integer = 0 To fieldNames.Length - 1
            lastValues(i) = sourceRow(fieldNames(i))
        Next
    End Sub




    'Read more: http://discuss.itacumens.com/index.php?topic=49368.0#ixzz32GD5qCYM

End Module
