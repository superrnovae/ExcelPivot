using FastMember;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace ConsoleApp1
{
    public static class ExcelHelper
    {
        private static readonly Dictionary<Type, Action<ICell, object>> cellFormatMapper = new()
        {
            { typeof(int), (ICell cell, object value) =>  cell.SetCellValue((int) value) },
            { typeof(double), (ICell cell, object value) => cell.SetCellValue((double) value) },
            { typeof(DateTime), (ICell cell, object value) => cell.SetCellValue((DateTime) value)},
            { typeof(string), (ICell cell, object value) => cell.SetCellValue((string) value) },
            { typeof(decimal), (ICell cell, object value) => cell.SetCellValue(decimal.ToDouble((decimal) value)) },
            { typeof(bool), (ICell cell, object value) => cell.SetCellValue((bool) value) },
            { typeof(byte), (ICell cell, object value) => cell.SetCellValue(Convert.ToDouble((byte) value)) },
            { typeof(ushort), (ICell cell, object value) => cell.SetCellValue(Convert.ToDouble((ushort) value)) },
            { typeof(short), (ICell cell, object value) => cell.SetCellValue(Convert.ToDouble((short) value)) },
            { typeof(long), (ICell cell, object value) => cell.SetCellValue(Convert.ToDouble((long) value)) },
            { typeof(char), (ICell cell, object value) => cell.SetCellValue(((char) value).ToString()) },
            { typeof(float), (ICell cell, object value) => cell.SetCellValue(Convert.ToDouble((float) value)) },
            { typeof(Guid), (ICell cell, object value) => cell.SetCellValue(((Guid) value).ToString()) },
            { typeof(object), (ICell cell, object value) => cell.SetCellValue(value.ToString()) }
        };

        public static void GenerateExcelFile<T>(Stream stream, IEnumerable<T> objects, PivotSettings settings)
        {
            if (objects == null) throw new ArgumentOutOfRangeException(nameof(objects));

            // creates objectreader, instead of reinventing the wheel with reflection
            using var reader = ObjectReader.Create(objects);

            // read column names from the dictionary
            var columns = new string[reader.FieldCount];

            // Store column lengths to manually adjust set later on
            Span<int> maxNumCharactersInColumns = columns.Length <= 1024 ? stackalloc int[columns.Length] : new int[columns.Length];

            for (int i = 0; i < reader.FieldCount; i++)
            {
                columns[i] = reader.GetName(i);
                int length = columns[i].Length;

                if (maxNumCharactersInColumns[i] < length)
                {
                    maxNumCharactersInColumns[i] = length + 4;
                }
            }

            // Create Worksheet
            using IWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("DATA");

            var table = InitTable(sheet, columns);


            // populate headings
            var headerRow = sheet.CreateRow(0);

            for (int i = 0; i < columns.Length; i++)
            {
                var cell = headerRow.CreateCell(i);
                cell.SetCellValue(columns[i]);
                cell.SetCellType(CellType.String);
            }

            // populate values
            ICell range_end = null;

            int index = 1;

            while (reader.Read())
            {
                var row = sheet.CreateRow(index);

                for (int j = 0; j < columns.Length; j++)
                {
                    ICell cell = row.CreateCell(j);
                    object value = reader[columns[j]];

                    //invoke method with respective type based on type of the column
                    cellFormatMapper[reader.GetFieldType(j)].Invoke(cell, value);

                    //compare lengths
                    int length = value != null ? value.ToString().Length : 0;

                    if (maxNumCharactersInColumns[j] < length)
                    {
                        maxNumCharactersInColumns[j] = length + 2;
                    }
                }

                range_end = row.GetCell(columns.Length - 1);

                index++;
            }

            /// Manually auto size columns, is way faster than sheet's AutoSizeColumn()
            for (int i = 0; i < columns.Length; i++)
            {
                int width = (int)(maxNumCharactersInColumns[i] * 1.30f) * 256; // 1.45f <- you can change this value
                sheet.SetColumnWidth(i, Math.Max(width, 2048)); // <- set calculated cell width
            }

            AreaReference myDataRange = new(new CellReference(0, 0), new CellReference(index - 1, columns.Length - 1));
            table.SetCellReferences(myDataRange);

            table.GetCTTable().autoFilter = new()
            {
                @ref = myDataRange.FormatAsString()
            };

            // Create dynamic table from existing one
            if (range_end != null && settings != null)
            {
                var pivotTable = table.Pivot(settings);
                sheet.IsSelected = false;
                workbook.SetActiveSheet(workbook.GetSheetIndex(pivotTable.GetParentSheet().SheetName));
            }

            // Write the workbook to file
            workbook.Write(stream, true);
        }

        private static XSSFTable InitTable(XSSFSheet sheet, string[] columns)
        {
            XSSFTable xssfTable = sheet.CreateTable();

            xssfTable.GetCTTable().id = 1;
            xssfTable.Name = "Data";
            xssfTable.IsHasTotalsRow = false;
            xssfTable.DisplayName = "MYTABLE";

            xssfTable.SetCellReferences(new AreaReference(new CellReference(0, 0), new CellReference(1, 1)));

            xssfTable.StyleName = XSSFBuiltinTableStyleEnum.TableStyleMedium16.ToString();
            xssfTable.Style.IsShowColumnStripes = false;
            xssfTable.Style.IsShowRowStripes = true;

            for (int i = 0; i < columns.Length; i++)
            {
                xssfTable.CreateColumn(columns[i], i);
            }

            return xssfTable;
        }
    }

    public class PivotSettings
    {
        public Dictionary<string, DataConsolidateFunction> ColumnLabels { get; set; }
        public string[] RowLabels { get; set; }
    }
}
