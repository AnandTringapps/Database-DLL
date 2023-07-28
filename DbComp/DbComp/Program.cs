using System.Data;
using OfficeOpenXml;
using System.Data.SqlClient;
using System.Diagnostics;

internal class Program
{
    private static void Main(string[] args)
    {
        string connectionString = "Data Source=LAPTOP-GMV7IJKD\\ANANDR;Initial Catalog=DataWarehouse;Integrated Security=True";
        string connectionString1 = "Data Source=LAPTOP-GMV7IJKD\\ANANDR;Initial Catalog=Staging;Integrated Security=True";

        // Retrieve the table names
        List<string> tableNames = GetTableNames(connectionString);
        List<string> tableNames1 = GetTableNames(connectionString1);

        List<string> alltable = tableNames.Intersect(tableNames1).ToList();

        // Find elements that are extra in List 1
        List<string> list1Extra = tableNames.Except(tableNames1).ToList();

        // Find elements that are extra in List 2
        List<string> list2Extra = tableNames1.Except(tableNames).ToList();

        WriteToExcel(alltable, list1Extra, list2Extra,connectionString,connectionString1);
    }
    static List<string> GetTableNames(string connectionString)
    {
        List<string> tableNames = new List<string>();

        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();

            // Get the schema information for the tables in the database
            DataTable schemaTable = connection.GetSchema("Tables");

            // Filter the schema information to retrieve only table names
            foreach (DataRow row in schemaTable.Rows)
            {
                string? tableName = row["TABLE_NAME"].ToString();
                tableNames.Add(tableName);
            }
        }

        return tableNames;
    }
    static void WriteToExcel(List<string> commonElements, List<string> list1Extra, List<string> list2Extra,string c1, string c2)
    {
        string filePath = "C:\\Users\\Anandth Ravichandran\\OneDrive\\Desktop\\";
        SqlConnectionStringBuilder b1 = new SqlConnectionStringBuilder(c1);
        SqlConnectionStringBuilder b2 = new SqlConnectionStringBuilder(c2);
        filePath =filePath + b1.InitialCatalog +"Vs"+ b2.InitialCatalog +".xlsx";

        string baseFileName = Path.GetFileNameWithoutExtension(filePath);
        string extension = Path.GetExtension(filePath);
        string folder = Path.GetDirectoryName(filePath);
        int increment = 0;
        string newFilePath = filePath;

        while (File.Exists(newFilePath))
        {
            increment++;
            newFilePath = Path.Combine(folder, $"{baseFileName}({increment}){extension}");
        }
        ExcelPackage.LicenseContext = LicenseContext.Commercial;
        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Comparison");

            // Write headers
            worksheet.Cells[1, 1].Value = "Common Tables";
            worksheet.Cells[1, 2].Value = "Database1 Extra";
            worksheet.Cells[1, 3].Value = "Database2 Extra";

            // Write data to the columns
            int row = 2;
            foreach (var commonElement in commonElements)
            {
                worksheet.Cells[row, 1].Value = commonElement;
                row++;
            }

            row = 2;
            foreach (var item in list1Extra)
            {
                worksheet.Cells[row, 2].Value = item;
                row++;
            }

            row = 2;
            foreach (var item in list2Extra)
            {
                worksheet.Cells[row, 3].Value = item;
                row++;
            }
            worksheet.Cells.AutoFitColumns();
            // Save the Excel file
            FileInfo excelFile = new FileInfo(newFilePath);
            package.SaveAs(excelFile);

        }
        Process.Start(new ProcessStartInfo(newFilePath)
        {
            UseShellExecute = true,
            Verb = "new"
        });
        Console.WriteLine("Comparison result has been written to ComparisonResult.xlsx");
    }
}