
using OfficeOpenXml;
using System.Drawing.Text;
namespace Excel_sheet_Import_and_Export
{
    class Program
    {
        static void Main(string[] args)
        {
            string file = @"C:\Users\Tenece\Desktop\data.xlsx";
            using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var sheet = package.Workbook.Worksheets[0];
                var persons = new Program().GetList<Person>(sheet);
                foreach (Person person in persons)
                {
                    Console.WriteLine($"{person.FirstName} \t {person.LastName} \t {person.Sex} \t {person.Mobile} \t {person.State}");
                }
            }
        }

        private List<T> GetList<T>(ExcelWorksheet sheet)
        {
            List<T> list = new List<T>();
            var columnInfo = Enumerable.Range(1, sheet.Dimension.Columns).ToList().Select(n =>
                
                new { Index = n, ColumnName = sheet.Cells[1, n].Value.ToString()}
            );
            for(int row =2; row < sheet.Dimension.Rows; row++)
            {
                T obj = (T)Activator.CreateInstance(typeof(T)); //generic object
                foreach(var prop in typeof(T).GetProperties())
                {
                    int col = columnInfo.SingleOrDefault(c => c.ColumnName == prop.Name).Index;
                    var val = sheet.Cells[row, col].Value;
                    var propType = prop.PropertyType;
                    prop.SetValue(obj, Convert.ChangeType(val, propType));
                }
                list.Add(obj);
            }
            return list;
        }
    }
}


