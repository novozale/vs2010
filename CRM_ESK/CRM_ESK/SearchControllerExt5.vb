Imports Syncfusion.WinForms.DataGrid
Imports Syncfusion.WinForms.DataGrid.Styles

Public Class SearchControllerExt5
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
        '// sfdatagrid7
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MyBase.HighlightSearchText(paint, column, style, bounds, cellValue, rowColumnIndex)
    End Sub
End Class
