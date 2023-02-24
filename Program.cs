using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using FastMember;
using NPOI.OpenXmlFormats.Spreadsheet;

namespace ConsoleApp1
{
    internal class Program
    {
        private static readonly Dictionary<Type, Action<ICell, object>> formatMapper = new()
        {
            { typeof(int), (ICell cell, object value) => cell.SetCellValue((int) value) },
            { typeof(double), (ICell cell, object value) => cell.SetCellValue((double) value) },
            { typeof(DateTime), (ICell cell, object value) => cell.SetCellValue((DateTime) value) },
            { typeof(string), (ICell cell, object value) => cell.SetCellValue(value.ToString()) },
            { typeof(decimal), (ICell cell, object value) => cell.SetCellValue(decimal.ToDouble((decimal) value)) },
            { typeof(bool), (ICell cell, object value) => cell.SetCellValue((bool) value) },
            { typeof(byte), (ICell cell, object value) => cell.SetCellValue((byte) value) }
        };

        
        static void Main(string[] args)
        {
            var settings = new PivotSettings()
            {
                ColumnLabels = new Dictionary<string, DataConsolidateFunction>()
                {
                    { "CA_NET", DataConsolidateFunction.SUM },
                    { "CA_BRUT", DataConsolidateFunction.SUM },
                    { "QTE_VENDUE", DataConsolidateFunction.SUM }
                },
                RowLabels = new string[] { "PRODUCT", "MONTH" }
            };

            GenerateExcelFile(GetData(), settings);
        }

        private static void GenerateExcelFile<T>(IEnumerable<T> objects, PivotSettings settings)
        {
            if (objects == null) throw new ArgumentOutOfRangeException(nameof(objects));

            // creates objectreader, instead of reinventing the wheel with reflection
            using var reader = ObjectReader.Create(objects);

            // store column names and their types in a dictionary
            var dict = new Dictionary<string, Type>();

            for(int i=0; i<reader.FieldCount; i++)
                dict.Add(reader.GetName(i), reader.GetFieldType(i));

            // Create Worksheet
            using IWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("DATA");

            string[] headings = dict.Keys.ToArray();
            int[] maxNumCharactersInColumns = new int[headings.Length];

            // populate headings
            var headerRow = sheet.CreateRow(0);

            for (int i = 0; i < headings.Length; i++)
            {
                var cell = headerRow.CreateCell(i);
                cell.SetCellValue(headings[i]);
                cell.SetCellType(CellType.String);


                // store objects length for later use in column width adjustment 
                int length = headings[i].Length;

                if (maxNumCharactersInColumns[i] < length)
                    maxNumCharactersInColumns[i] = length + 4;

            }

            // populate values
            ICell range_end = null;

            int index = 1;

            while(reader.Read())
            {
                var row = sheet.CreateRow(index);

                for (int j=0; j<dict.Keys.Count; j++)
                {
                    ICell cell = row.CreateCell(j);
                    Type type = dict.Values.ElementAt(j);
                    object value = reader[headings[j]];

                    //invoke method with respective type based on type of the column
                    formatMapper[type].Invoke(cell, value);

                    //compare lengths
                    int length = value != null ? value.ToString().Length : 0;

                    if (maxNumCharactersInColumns[j] < length)
                    { // adjust the columns width
                        maxNumCharactersInColumns[j] = length+2;
                    }
                }

                range_end = row.GetCell(headings.Length-1);

                index++;
            }

            // Manually auto size columns
            for (int i = 0; i < headings.Length; i++)
            {
                int width = (int)(maxNumCharactersInColumns[i] * 1.25f) * 256; // 1.45f <- you can change this value
                sheet.SetColumnWidth(i, Math.Max(width, 2048)); // <- set calculated cell width
            }

            AreaReference myDataRange = new(new CellReference(0, 0), new CellReference(index - 1, headings.Length - 1));

            // Format existing data as table
            var table = FormatAsTable(sheet, myDataRange, headings);

            // Create dynamic table from existing one
            if (range_end != null && settings != null)
            {
                var pivotTable = CreatePivotTable(workbook, sheet, table, settings);
                sheet.IsSelected = false;
                workbook.SetActiveSheet(workbook.GetSheetIndex(pivotTable.GetParentSheet().SheetName));
            }

            // Write the workbook to file
            using var fileWritter = new FileStream("test.xlsx", FileMode.Create, FileAccess.Write);
            workbook.Write(fileWritter, false);
        }

        private static XSSFTable FormatAsTable(XSSFSheet sheet, AreaReference range, string[] columns)
        {
            // Format existing data as table
            XSSFTable xssfTable = sheet.CreateTable();
            CT_Table ctTable = xssfTable.GetCTTable();
            
            ctTable.@ref = range.FormatAsString();
            ctTable.id = 1;
            ctTable.name = "Table1";
            ctTable.displayName = "Table1";

            ctTable.tableStyleInfo = new()
            {
                name = "TableStyleMedium2", // TableStyleMedium2 is one of XSSFBuiltinTableStyle
                showRowStripes = true

            };

            ctTable.totalsRowShown = false;
            ctTable.autoFilter = new()
            {
                @ref = range.FormatAsString()
            };

            ctTable.tableColumns = new CT_TableColumns
            {
                tableColumn = new List<CT_TableColumn>()
            };

            for (int i = 0; i < columns.Length; i++)
                ctTable.tableColumns.tableColumn.Add(new CT_TableColumn() { id = (uint)i + 1, name = columns[i] });

            return xssfTable;
        }


        private static XSSFPivotTable CreatePivotTable(IWorkbook workbook, XSSFSheet sheet, XSSFTable table, PivotSettings settings)
        {
            var startReference = new CellReference(table.StartRowIndex, table.StartColIndex);
            var endReference = new CellReference(table.EndRowIndex, table.EndColIndex);
            var range = new AreaReference(startReference, endReference);

            XSSFSheet pivotSheet = (XSSFSheet)workbook.CreateSheet("PIVOT");
            XSSFPivotTable pivotTable = pivotSheet.CreatePivotTable(range, new CellReference(0, 0), sheet);

            foreach(var item in settings.RowLabels)
            {
                int ix = table.FindColumnIndex(item);

                if(ix != -1)
                    pivotTable.AddRowLabel(ix);
            }

            foreach(var item in settings.ColumnLabels.Keys)
            {
                int ix = table.FindColumnIndex(item);

                if (ix != -1)
                    pivotTable.AddColumnLabel(settings.ColumnLabels[item], ix, item);
            }

            return pivotTable;
        }


        private static IEnumerable<Stats> GetData()
        {
            for(int j=0; j<10000; j++)
            {
                var obj = new Stats()
                {
                    PRODUCT = "TEST" + j.ToString(),
                };

                for (int i = 1; i <= 12; i++)
                {
                    obj.MONTH = "2022-" + i.ToString();
                    obj.CA_NET = new Random().Next(10000);
                    obj.CA_BRUT = new Random().Next(10000);
                    obj.QTE_VENDUE = new Random().Next(100);

                    yield return obj;
                }
            }
        }
    }

    public class PivotSettings
    {
        public Dictionary<string, DataConsolidateFunction> ColumnLabels { get; set; }
        public string[] RowLabels { get; set; }
    }


    internal class Stats
    {
        public string PRODUCT { get; set; }
        public string MONTH { get; set; }
        public int CA_NET { get; set; }
        public int CA_BRUT { get; set;}
        public int QTE_VENDUE { get; set; }
    }
}

