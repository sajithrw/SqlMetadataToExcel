using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data.SqlClient;

if (args.Length < 2)
{
    Console.WriteLine("Usage: dotnet run <connectionString> <outputFilePath>");
    return;
}

string connectionString = args[0];
string outputPath = args[1];

List<TableMetadata> tables = GetDatabaseMetadata(connectionString);
WriteMetadataToExcel(tables, outputPath);

static List<TableMetadata> GetDatabaseMetadata(string connectionString)
{
    var tables = new List<TableMetadata>();

    using (SqlConnection conn = new SqlConnection(connectionString))
    {
        conn.Open();

        // Get all table names
        using (SqlCommand cmd = new SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'", conn))
        {
            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    string tableName = reader.GetString(0);
                    var table = new TableMetadata { TableName = tableName, Columns = new List<ColumnMetadata>() };

                    // Get columns for each table
                    using (SqlCommand colCmd = new SqlCommand(@"
                            SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE, COLUMNPROPERTY(object_id(TABLE_NAME), COLUMN_NAME, 'IsIdentity') AS IsIdentity
                            FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @tableName", conn))
                    {
                        colCmd.Parameters.AddWithValue("@tableName", tableName);
                        using (SqlDataReader colReader = colCmd.ExecuteReader())
                        {
                            while (colReader.Read())
                            {
                                //var dataType = colReader.GetString(1);
                                //var maxLength = colReader.IsDBNull(2) ? (int?)null : colReader.GetInt32(2);
                                //var numericPrecision = colReader.IsDBNull(3) ? (byte?)null : colReader.GetByte(3);
                                //var numericScale = colReader.IsDBNull(4) ? (int?)null : colReader.GetInt32(4);
                                //var fullDataType = GetFullDataType(dataType, maxLength, numericPrecision, numericScale);

                                table.Columns.Add(new ColumnMetadata
                                {
                                    ColumnName = colReader.GetString(0),
                                    DataType = colReader.GetString(1),
                                    IsNullable = colReader.GetString(2) == "YES",
                                    IsPrimaryKey = false,
                                    IsForeignKey = false
                                });
                            }
                        }
                    }

                    // Get primary key columns for each table
                    using (SqlCommand keyCmd = new SqlCommand(@"
                            SELECT COLUMN_NAME
                            FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE
                            WHERE TABLE_NAME = @tableName AND CONSTRAINT_NAME LIKE 'PK_%'", conn))
                    {
                        keyCmd.Parameters.AddWithValue("@tableName", tableName);
                        using (SqlDataReader keyReader = keyCmd.ExecuteReader())
                        {
                            while (keyReader.Read())
                            {
                                string keyColumn = keyReader.GetString(0);
                                var column = table.Columns.Find(c => c.ColumnName == keyColumn);
                                if (column != null)
                                {
                                    column.IsPrimaryKey = true;
                                }
                            }
                        }
                    }

                    // Get foreign key columns for each table
                    using (SqlCommand fkCmd = new SqlCommand(@"
                    SELECT CU.COLUMN_NAME
                    FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC
                    INNER JOIN INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE CU ON CU.CONSTRAINT_NAME = RC.CONSTRAINT_NAME
                    WHERE CU.TABLE_NAME = @tableName", conn))
                    {
                        fkCmd.Parameters.AddWithValue("@tableName", tableName);
                        using (SqlDataReader fkReader = fkCmd.ExecuteReader())
                        {
                            while (fkReader.Read())
                            {
                                string fkColumn = fkReader.GetString(0);
                                var column = table.Columns.Find(c => c.ColumnName == fkColumn);
                                if (column != null)
                                {
                                    column.IsForeignKey = true;
                                }
                            }
                        }
                    }

                    tables.Add(table);
                }
            }
        }
    }

    return tables;
}

static string GetFullDataType(string dataType, int? maxLength, byte? numericPrecision, int? numericScale)
{
    if (dataType == "varchar" || dataType == "char" || dataType == "nvarchar" || dataType == "nchar")
    {
        return maxLength.HasValue ? $"{dataType}({maxLength})" : dataType;
    }
    else if (dataType == "decimal" || dataType == "numeric")
    {
        return numericPrecision.HasValue && numericScale.HasValue ? $"{dataType}({numericPrecision},{numericScale})" : dataType;
    }
    else
    {
        return dataType;
    }
}

static void WriteMetadataToExcel(List<TableMetadata> tables, string outputPath)
{
    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

    using (ExcelPackage package = new ExcelPackage())
    {
        foreach (var table in tables)
        {
            var worksheet = package.Workbook.Worksheets.Add(table.TableName);

            worksheet.Cells[1, 1].Value = "Column Name";
            worksheet.Cells[1, 2].Value = "Data Type";
            worksheet.Cells[1, 3].Value = "Is Nullable";
            worksheet.Cells[1, 4].Value = "Is Primary Key";
            worksheet.Cells[1, 5].Value = "Is Foreign Key";

            for (int i = 0; i < table.Columns.Count; i++)
            {
                var column = table.Columns[i];
                worksheet.Cells[i + 2, 1].Value = column.ColumnName;
                worksheet.Cells[i + 2, 2].Value = column.DataType;
                worksheet.Cells[i + 2, 3].Value = column.IsNullable;
                worksheet.Cells[i + 2, 4].Value = column.IsPrimaryKey;
                worksheet.Cells[i + 2, 5].Value = column.IsForeignKey;
            }

            // Add borders to all cells
            using (var range = worksheet.Cells[1, 1, table.Columns.Count + 1, 5])
            {
                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            // Auto-fit columns
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
        }

        package.SaveAs(new FileInfo(outputPath));
    }
}

class TableMetadata
{
    public string TableName { get; set; }
    public List<ColumnMetadata> Columns { get; set; }
}

class ColumnMetadata
{
    public string ColumnName { get; set; }
    public string DataType { get; set; }
    public bool IsNullable { get; set; }
    public bool IsPrimaryKey { get; set; }
    public bool IsForeignKey { get; set; }
}