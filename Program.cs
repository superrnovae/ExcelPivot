using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using FastMember;

namespace ConsoleApp1
{
    internal class Program
    {
        

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

            using var stream = File.Open("test.xlsx", FileMode.Create, FileAccess.Write);

            ExcelHelper.GenerateExcelFile(stream, GenerateObjects(), settings);
        }


        private static IEnumerable<Stats> GenerateObjects()
        {
            for(int j=0; j<2000; j++)
            {
                var obj = new Stats()
                {
                    PRODUCT = "TEST" + j.ToString(),
                    GUID = Guid.NewGuid()
                };

                for (int i = 1; i <= 12; i++)
                {
                    obj.MONTH = "2022-" + i.ToString("00");
                    obj.CA_NET = new Random().Next(100);
                    obj.CA_BRUT = new Random().Next(200);
                    obj.QTE_VENDUE = new Random().Next(100);

                    yield return obj;
                }
            }
        }
    }

    internal class Stats
    {
        public string PRODUCT { get; set; }
        public string MONTH { get; set; }
        public int CA_NET { get; set; }
        public int CA_BRUT { get; set;}
        public int QTE_VENDUE { get; set; }

        public Guid GUID { get; set; }

    }
}

