'Public Class cDataGridViewPaintCellBorder_OLD
'    Inherits DataGridViewTextBoxCell

'    Protected Overrides Sub Paint( _
'    ByVal graphics As Graphics, _
'    ByVal clipBounds As Rectangle, _
'    ByVal cellBounds As Rectangle, _
'    ByVal rowIndex As Integer, _
'    ByVal elementState As DataGridViewElementStates, _
'    ByVal value As Object, _
'    ByVal formattedValue As Object, _
'    ByVal errorText As String, _
'    ByVal cellStyle As DataGridViewCellStyle, _
'    ByVal advancedBorderStyle As DataGridViewAdvancedBorderStyle, _
'    ByVal paintParts As DataGridViewPaintParts)

'        ' Call the base class method to paint the default cell appearance.
'        MyBase.Paint(graphics, clipBounds, cellBounds, rowIndex, elementState, _
'            value, formattedValue, errorText, cellStyle, _
'            advancedBorderStyle, paintParts)

'        ' Retrieve the client location of the mouse pointer.
'        Dim cursorPosition As Point = _
'            Me.DataGridView.PointToClient(Cursor.Position)

'        ' If the mouse pointer is over the current cell, draw a custom border.
'        If cellBounds.Contains(cursorPosition) Then
'            Dim newRect As New Rectangle(cellBounds.X + 1, _
'                cellBounds.Y + 1, cellBounds.Width - 4, _
'                cellBounds.Height - 4)
'            graphics.DrawRectangle(Pens.Red, newRect)
'        End If

'    End Sub

'End Class
