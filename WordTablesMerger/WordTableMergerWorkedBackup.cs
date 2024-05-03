//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Wordprocessing;
//using Serilog;
//using System.Text;
//using System.Text.RegularExpressions;

//namespace WordToExcelMigrator
//{
//    /**
//     * IMPORTANT: Merged document includes header and first row from the first table. 
//     * However, it doesn't take header and the first row from all other tables (documents).
//     */
//    public class WordTableMergerWorkedBackup
//    {
//        private static string _sourceDirectory;
//        private static List<string> _sourcePaths;

//        private static string _destinationPath;

//        static void Main(string[] args)
//        {
//            InitializeLogger();
//            InitializeAndValidateFiles();

//            WordTableMergerWorkedBackup merger = new();
//            merger.MergeDocxTables();

//            Log.CloseAndFlush();
//        }

//        private static void InitializeLogger()
//        {
//            Log.Logger = new LoggerConfiguration()
//              .WriteTo.Console()
//              .CreateLogger();
//        }

//        private static void InitializeAndValidateFiles()
//        {
//            _sourceDirectory = @"C:\Projects\Pipi.Automate\filesblabla\source2";
//            if (!Directory.Exists(_sourceDirectory))
//            {
//                Log.Error("The source directory does not exist: {0}", _sourceDirectory);
//                Environment.Exit(1);
//            }

//            _sourcePaths = Directory.GetFiles(_sourceDirectory, "*.docx")
//                .OrderBy(path => path)
//                .ToList();
//            _destinationPath = Path.Combine(_sourceDirectory, "merged-file.docx");

//            Log.Information("Initialization complete. Source directory: {0}", _sourceDirectory);
//        }

//        private void MergeDocxTables()
//        {
//            Log.Information(Common.PrintPipi());
//            Log.Information("Starting merge process.");

//            if (!_sourcePaths.Any())
//            {
//                Log.Error("No .docx files found in the source directory. Process terminated.");
//                return;
//            }

//            try
//            {
//                using (var destinationDoc = WordprocessingDocument.Create(_destinationPath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
//                {
//                    var mainPart = destinationDoc.AddMainDocumentPart();
//                    mainPart.Document = new Document(new Body());
//                    Table mergedTable = new Table();
//                    bool isFirstTable = true;
//                    bool isFirstFile = true;

//                    foreach (var sourcePath in _sourcePaths)
//                    {
//                        using (var sourceDoc = WordprocessingDocument.Open(sourcePath, false))
//                        {
//                            var tables = sourceDoc.MainDocumentPart.Document.Body.Elements<Table>().ToList();

//                            if (!tables.Any())
//                            {
//                                Log.Information($"No tables found in document {sourcePath}. Skipping.");
//                                continue;
//                            }

//                            bool isFirstTableInFile = true;
//                            for (int i = 0; i < tables.Count; i++)
//                            {
//                                if (isFirstTable)
//                                {
//                                    mergedTable = (Table)tables[i].CloneNode(true);

//                                    var rows = mergedTable.Elements<TableRow>().ToList();
//                                    for (int j = 0; j < rows.Count; j++)
//                                    {
//                                        ProcessRowCells(rows[j], isFirstFile && j == 0);
//                                    }
//                                    isFirstTable = false;
//                                    isFirstFile = false;
//                                }
//                                else
//                                {
//                                    foreach (var row in tables[i].Elements<TableRow>().Skip(2))
//                                    {
//                                        var clonedRow = (TableRow)row.CloneNode(true);
//                                        ProcessRowCells(clonedRow, isFirstTableInFile);
//                                        mergedTable.AppendChild(clonedRow);
//                                    }
//                                }
//                                isFirstTableInFile = false;
//                            }
//                        }
//                        Log.Information($"Successfully processed document {sourcePath}");
//                    }

//                    RemoveNinthColumnFromMergedTable(mergedTable);
//                    mainPart.Document.Body.AppendChild(mergedTable);
//                    mainPart.Document.Save();
//                    Log.Information("Tables merged successfully.");
//                }
//            }
//            catch (Exception ex)
//            {
//                Log.Error(ex, "An error occurred during the merge process.");
//                throw;
//            }
//        }

//        private void ValidateSourceFiles()
//        {
//            if (!_sourcePaths.Any())
//            {
//                Log.Error("No .docx files found in the source directory. Process terminated.");
//                return;
//            }
//        }

//        private void ProcessRowCells(TableRow row, bool isFirstRowInNewFile = false)
//        {
//            var cells = row.Elements<TableCell>().ToList();
//            for (int i = 0; i < cells.Count; i++)
//            {
//                if (i == 9)
//                {
//                    RemoveNewLineIn10thColumnForSingleCell(cells[i]);
//                }
//                else
//                {
//                    RemoveNewLineForSingleCell(cells[i]);
//                }
//            }

//            //If first row for a first table from new file, mark the whole row as yellow.
//            if (isFirstRowInNewFile)
//            {
//                foreach (var cell in cells)
//                {
//                    // Ensure TableCellProperties exists for shading
//                    if (cell.Elements<TableCellProperties>().Any())
//                    {
//                        var cellProperties = cell.Elements<TableCellProperties>().First();
//                        cellProperties.Shading = new Shading() { Fill = "FFFF00" };
//                    }
//                    // If no TableCellProperties, create and append it
//                    else
//                    {
//                        cell.Append(new TableCellProperties(new Shading() { Fill = "FFFF00" }));
//                    }
//                }
//            }
//        }

//        private void RemoveNewLineForSingleCell(TableCell cell)
//        {
//            var cleanedText = GetCleanedTextFromCell(cell);

//            cell.RemoveAllChildren<Paragraph>();
//            var newParagraph = new Paragraph(new Run(new Text(cleanedText)));
//            cell.Append(newParagraph);
//        }

//        private void RemoveNewLineIn10thColumnForSingleCell(TableCell cell)
//        {
//            var cleanedText = GetCleanedTextFromCell(cell);

//            // Apply specific modifications for the 10th column.
//            var pattern = @"^\d{4}\sPL\d+";
//            var match = Regex.Match(cleanedText, pattern);
//            var newText = match.Success ? match.Value : cleanedText;

//            cell.RemoveAllChildren<Paragraph>();
//            var newParagraph = new Paragraph(new Run(new Text(newText)));
//            cell.AppendChild(newParagraph);
//        }

//        private string GetCleanedTextFromCell(TableCell cell)
//        {
//            var cellText = new StringBuilder();
//            var paragraphs = cell.Elements<Paragraph>();

//            foreach (var para in paragraphs)
//            {
//                foreach (var run in para.Elements<Run>())
//                {
//                    foreach (var text in run.Elements<Text>())
//                    {
//                        cellText.Append(text.Text.Replace("\r", "").Replace("\n", " "));
//                    }
//                }
//            }

//            return cellText.ToString();
//        }

//        private void RemoveNinthColumnFromMergedTable(Table table)
//        {
//            foreach (var row in table.Elements<TableRow>())
//            {
//                var cells = row.Elements<TableCell>().ToList();
//                if (cells.Count >= 9)
//                {
//                    cells[8].Remove();
//                }
//            }
//        }

//        //private void MergeDocxTables2()
//        //{
//        //    Log.Information(Common.PrintPipi());
//        //    Log.Information("Starting merge process.");

//        //    try
//        //    {
//        //        using (var destinationDoc = WordprocessingDocument.Create(_destinationPath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
//        //        {
//        //            var mainPart = destinationDoc.AddMainDocumentPart();
//        //            mainPart.Document = new Document(new Body());

//        //            Table mergedTable = null;

//        //            //Iterate each file
//        //            for (int i = 0; i < _filePaths.Count; i++)
//        //            {
//        //                try
//        //                {
//        //                    //Open file
//        //                    using (var doc = WordprocessingDocument.Open(_filePaths[i], false))
//        //                    {
//        //                        var docBody = doc.MainDocumentPart.Document.Body;
//        //                        var table = docBody.Elements<Table>().FirstOrDefault();

//        //                        if (table == null)
//        //                        {
//        //                            Log.Warning("No table found in document {DocumentPath}. Skipping.", _filePaths[i]);
//        //                            continue;
//        //                        }

//        //                        //Process first file
//        //                        if (i == 0)
//        //                        {
//        //                            mergedTable = (Table)table.CloneNode(true);
//        //                            RemoveNewLineIn10thColumn(mergedTable);
//        //                            mainPart.Document.Body.AppendChild(mergedTable);
//        //                        }
//        //                        //Process next files
//        //                        else
//        //                        {
//        //                            foreach (var row in table.Elements<TableRow>().Skip(2))
//        //                            {
//        //                                var clonedRow = (TableRow)row.CloneNode(true);
//        //                                RemoveNewLineIn10thColumnForSingleRow(clonedRow);
//        //                                mergedTable.AppendChild(clonedRow);
//        //                            }
//        //                        }
//        //                        Log.Information($"Successfully processed document {_filePaths[i]} | {i}/{_filePaths.Count}");
//        //                    }
//        //                }
//        //                catch (Exception ex)
//        //                {
//        //                    Log.Error(ex, $"Failed to process document {_filePaths[i]}.");
//        //                }
//        //            }

//        //            RemoveNinthColumnFromMergedTable(mergedTable);
//        //            mainPart.Document.Save();

//        //            Log.Information("Documents merged successfully.");
//        //        }
//        //    }
//        //    catch (Exception ex)
//        //    {
//        //        Log.Error(ex, "An error occurred during the merge process.");
//        //        throw;
//        //    }
//        //}
//    }
//}