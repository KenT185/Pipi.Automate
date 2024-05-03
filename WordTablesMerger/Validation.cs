using DocumentFormat.OpenXml.Packaging;
using Serilog;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordToExcelMigrator
{
    public static class Validation
    {
        private static int CountExpectedRows(List<Table> tables)
        {
            int rowCount = 0;
            for (int i = 0; i < tables.Count; i++)
            {
                // Add all rows from the first table
                // For subsequent tables, skip header and first row
                rowCount += (i == 0) ? tables[i].Elements<TableRow>().Count() : tables[i].Elements<TableRow>().Count() - 2;
            }
            return rowCount;
        }

        private static int CountActualRows(string filePath)
        {
            using (var doc = WordprocessingDocument.Open(filePath, false))
            {
                var table = doc.MainDocumentPart.Document.Body.Elements<Table>().FirstOrDefault();
                if (table != null)
                {
                    return table.Elements<TableRow>().Count();
                }
            }
            return 0;
        }

        private static void ValidateDocumentRowMatch(int expectedRows, int actualRows)
        {
            if (expectedRows == actualRows)
            {
                Log.Information($"Validation passed. Both original and merged documents have the expected number of rows: {actualRows}.");
            }
            else
            {
                Log.Information($"Validation failed. Expected rows: {expectedRows}, but found: {actualRows} in the merged document.");
            }
        }
    }
}