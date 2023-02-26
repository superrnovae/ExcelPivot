using ConsoleApp1;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

public static class ProgramHelpers
{

    public static XSSFPivotTable Pivot(this XSSFTable table, PivotSettings settings)
    {
        var startReference = new CellReference(table.StartRowIndex, table.StartColIndex);
        var endReference = new CellReference(table.EndRowIndex, table.EndColIndex);
        var range = new AreaReference(startReference, endReference);

        var pivotTablePosition = new CellReference(0, 0);

        var sourceSheet = table.GetXSSFSheet();
        var workbook = sourceSheet.Workbook;

        var pivotSheet = workbook.CreateSheet() as XSSFSheet;

        XSSFPivotTable pivotTable = pivotSheet.CreatePivotTable(range, new CellReference(0, 0), sourceSheet);

        foreach (var item in settings.RowLabels)
        {
            int ix = table.FindColumnIndex(item);

            if (ix == -1) continue;

            pivotTable.AddRowLabel(ix);
        }

        foreach (var item in settings.ColumnLabels.Keys)
        {
            int ix = table.FindColumnIndex(item);

            if (ix == -1) continue;

            pivotTable.AddColumnLabel(settings.ColumnLabels[item], ix, item);
        }

        return pivotTable;
    }
}