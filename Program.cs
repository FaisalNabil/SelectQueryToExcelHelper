using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using OfficeOpenXml;
using System.Windows.Forms;

public class Program
{
    [STAThread] // Required for OpenFileDialog to work properly
    static void Main()
    {
        try
        {
            // Set the EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Step 1: Read DB connection string from a text file
            string connString = File.ReadAllText("db_connection.txt").Trim();

            // Step 2: Let the user select the SQL file using a file picker
            string sqlFilePath = GetSQLFilePath();
            if (string.IsNullOrEmpty(sqlFilePath))
            {
                Console.WriteLine("No file selected. Exiting...");
                return;
            }

            string[] queries = File.ReadAllText(sqlFilePath).Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);

            // Step 3: Execute each SELECT query and write results to Excel
            string excelFilePath = Path.ChangeExtension(sqlFilePath, ".xlsx");

            using (SqlConnection conn = new SqlConnection(connString))
            using (ExcelPackage excel = new ExcelPackage())
            {
                conn.Open();

                int sheetIndex = 1;
                foreach (string query in queries)
                {
                    string cleanQuery = query.Trim();
                    if (string.IsNullOrWhiteSpace(cleanQuery)) continue;

                    DataTable table = ExecuteQuery(conn, cleanQuery);
                    AddSheetToExcel(excel, table, "Sheet" + sheetIndex, cleanQuery);
                    sheetIndex++;
                }

                // Save Excel file
                File.WriteAllBytes(excelFilePath, excel.GetAsByteArray());
            }

            Console.WriteLine($"Excel file created successfully: {excelFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }

        // ✅ Keeps CMD open after execution
        Console.WriteLine("\nProcess completed. Press any key to exit...");
        Console.ReadKey();
    }

    /// <summary>
    /// Opens a file dialog to let the user select a .sql file
    /// </summary>
    static string GetSQLFilePath()
    {
        Console.WriteLine("Drag and drop the SQL file here, or enter its full path:");
        string path = Console.ReadLine().Trim('"'); // Remove surrounding quotes if dragged

        if (File.Exists(path))
        {
            return path;
        }

        Console.WriteLine("Invalid file path. Try again.");
        return GetSQLFilePath(); // Recursive call to reattempt
    }

    static DataTable ExecuteQuery(SqlConnection conn, string query)
    {
        using (SqlCommand cmd = new SqlCommand(query, conn))
        using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
        {
            DataTable table = new DataTable();
            adapter.Fill(table);
            return table;
        }
    }

    static void AddSheetToExcel(ExcelPackage excel, DataTable table, string sheetName, string query)
    {
        var worksheet = excel.Workbook.Worksheets.Add(sheetName);

        int currentRow = 1;

        if (table.Columns.Count == 0)
        {
            // ✅ If the table has no columns, ensure at least a default message
            worksheet.Cells[currentRow, 1].Value = "No data returned from query.";
            currentRow += 2; // Leave a gap before adding the query
        }
        else
        {
            // ✅ Add headers even if no data is returned
            for (int col = 0; col < table.Columns.Count; col++)
            {
                worksheet.Cells[currentRow, col + 1].Value = table.Columns[col].ColumnName;
            }

            currentRow++;

            // ✅ Add data rows if available
            for (int row = 0; row < table.Rows.Count; row++)
            {
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    worksheet.Cells[currentRow, col + 1].Value = table.Rows[row][col];
                }
                currentRow++;
            }

            // ✅ If no data rows exist, leave a message
            if (table.Rows.Count == 0)
            {
                worksheet.Cells[currentRow, 1].Value = "(No records found)";
                currentRow++;
            }
        }

        // ✅ Leave a gap, then print the query at the bottom
        currentRow += 2;
        worksheet.Cells[currentRow, 1].Value = "Executed Query:";
        worksheet.Cells[currentRow + 1, 1].Value = query;

        // ✅ Format the query to wrap text and span multiple columns
        worksheet.Cells[currentRow + 1, 1, currentRow + 1, table.Columns.Count > 0 ? table.Columns.Count : 2].Merge = true;
        worksheet.Cells[currentRow + 1, 1].Style.WrapText = true;
    }
}
