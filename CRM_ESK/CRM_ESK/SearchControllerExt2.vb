Imports Syncfusion.WinForms.DataGrid
Imports Syncfusion.WinForms.DataGrid.Styles

Public Class SearchControllerExt2
    Inherits SearchController

    Public Sub New(ByVal grid As SfDataGrid)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Конструктор
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MyBase.New(grid)
    End Sub

    Protected Overrides Sub HighlightSearchText(ByVal paint As Graphics, ByVal column As DataColumnBase, ByVal style As CellStyleInfo, _
        ByVal bounds As Rectangle, ByVal cellValue As String, ByVal rowColumnIndex As Syncfusion.WinForms.GridCommon.ScrollAxis.RowColumnIndex)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// sfdatagrid2
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If column.GridColumn.MappingName = "CompanyName" Or _
            column.GridColumn.MappingName = "CompanyScalaCode" Or _
            column.GridColumn.MappingName = "CustProject" Or _
            column.GridColumn.MappingName = "OrdersQTY" Or _
            column.GridColumn.MappingName = "Status" Then
            MyBase.HighlightSearchText(paint, column, style, bounds, cellValue, rowColumnIndex)
        End If
    End Sub
End Class