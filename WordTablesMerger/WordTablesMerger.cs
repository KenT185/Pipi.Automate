using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Serilog;
using System.Text;
using System.Text.RegularExpressions;

namespace WordTablesMerger
{
    public class WordTablesMerger
    {
        private static string _sourceDirectory;
        private static List<string> _sourceFilePaths;
        private static string _targetFilePath;

        private static int _slaughterColumnIndex = 8;

        static void Main(string[] args)
        {
            InitializeLogger();
            InitializeAndValidateFiles();

            WordTablesMerger merger = new();
            merger.MergeDocxTables();

            Log.CloseAndFlush();
        }

        private static void InitializeLogger()
        {
            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .CreateLogger();
        }

        private static void InitializeAndValidateFiles()
        {
            Console.WriteLine("You will be prompted to enter the directory path where .docx files are placed. Keep there only files which you want to merge!");

            bool isValidDirectory = false;
            while (!isValidDirectory)
            {
                Console.Write("Enter directory path: ");
                _sourceDirectory = Console.ReadLine();

                if (!Directory.Exists(_sourceDirectory))
                {
                    Log.Error($"Given source directory {_sourceDirectory} does not exist. Try again.");
                    continue;
                }

                isValidDirectory = true;
            }

            _sourceFilePaths = Directory.GetFiles(_sourceDirectory, "*.docx").OrderBy(path => path).ToList();
            if (!_sourceFilePaths.Any())
            {
                Log.Error("No .docx files found in the source directory. Process terminated. Press any key to exit.");
                Console.ReadLine();
                Environment.Exit(1);
            }

            _targetFilePath = Path.Combine(_sourceDirectory, "merged-file.docx");
            if (File.Exists(_targetFilePath))
            {
                Log.Error($"Target file {_targetFilePath} already exists. Please remove this file or move to another directory. Process terminated. Press any key to exit.");
                Console.ReadLine();
                Environment.Exit(1);
            }
        }


        private void MergeDocxTables()
        {
            Log.Information(Common.PrintPipi());
            Log.Information("Starting merge process.");

            try
            {
                using (var targetFile = WordprocessingDocument.Create(_targetFilePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = targetFile.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());

                    Table targetTable = new Table();
                    TableProperties tableProperties = CreateTableProperties();
                    targetTable.AppendChild(tableProperties);

                    // Iterating source files
                    foreach (var sourceFilePath in _sourceFilePaths)
                    {
                        using (var sourceFile = WordprocessingDocument.Open(sourceFilePath, false))
                        {
                            var sourceTables = sourceFile.MainDocumentPart.Document.Body.Elements<Table>().ToList();
                            if (!sourceTables.Any())
                            {
                                Log.Information($"No tables found in document {sourceFilePath}. Skipping.");
                                continue;
                            }
                            bool isFirstRowInFirstTableInFile = true;

                            // Iterating source tables
                            for (int i = 0; i < sourceTables.Count; i++)
                            {
                                // Iterating source rows and skipping header and first row
                                foreach (var sourceRow in sourceTables[i].Elements<TableRow>().Skip(2))
                                {
                                    var targetRow = (TableRow)sourceRow.CloneNode(true);
                                    ModifyTargetTableCell(targetRow, isFirstRowInFirstTableInFile);
                                    isFirstRowInFirstTableInFile = false;

                                    targetTable.AppendChild(targetRow);
                                }
                            }
                        }

                        Log.Information($"Successfully processed document {sourceFilePath}");
                    }

                    RemoveNinthColumnFromTargetTable(targetTable);
                    mainPart.Document.Body.AppendChild(targetTable);
                    mainPart.Document.Save();
                    Log.Information($"Tables merged successfully. File saved in {_targetFilePath}. \n Press any key to exit");
                    Console.ReadLine();
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "An unexpected error occurred during the merge process.");
                throw;
            }
        }

        private TableProperties CreateTableProperties()
        {
            return new TableProperties(
                new TableBorders(
                    new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                    new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                    new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                    new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                    new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                    new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 }
                )
            );
        }

        private void ModifyTargetTableCell(TableRow row, bool isFirstRowInFirstTableInFile = false)
        {
            var fontProperties = CreateFontProperties();
            var cells = row.Elements<TableCell>().ToList();

            for (int i = 0; i < cells.Count; i++)
            {
                RemoveNewLineForCell(cells[i], i, fontProperties);

                if (isFirstRowInFirstTableInFile)
                {
                    MarkRowAllCellsAsYellow(cells);
                }
            }
        }

        private RunProperties CreateFontProperties()
        {
            return new RunProperties(
                new RunFonts { Ascii = "Times New Roman" },
                new FontSize { Val = "16" }  // Font size 8 (Half-point measurement, so "16")
            );
        }

        private void RemoveNewLineForCell(TableCell cell, int cellNumber, RunProperties fontProperties)
        {
            var cleanedText = GetCleanedTextFromCell(cell);
            if (cellNumber == 9)
            {
                var pattern = @"^\d{4}\sPL\d+";
                var match = Regex.Match(cleanedText, pattern);
                cleanedText = match.Success ? match.Value : cleanedText;
            }

            cell.RemoveAllChildren<Paragraph>();

            var runPropertiesClone = (RunProperties)fontProperties.Clone();
            var run = new Run(runPropertiesClone, new Text(cleanedText));
            var newParagraph = new Paragraph(run);

            cell.AppendChild(newParagraph);
        }

        private string GetCleanedTextFromCell(TableCell cell)
        {
            var cellText = new StringBuilder();
            var paragraphs = cell.Elements<Paragraph>();

            foreach (var para in paragraphs)
            {
                foreach (var run in para.Elements<Run>())
                {
                    foreach (var text in run.Elements<Text>())
                    {
                        cellText.Append(text.Text.Replace("\r", "").Replace("\n", " "));
                    }
                }
            }

            return cellText.ToString();
        }

        private void MarkRowAllCellsAsYellow(List<TableCell>? cells)
        {
            foreach (var cell in cells)
            {
                if (cell.Elements<TableCellProperties>().Any())
                {
                    var cellProperties = cell.Elements<TableCellProperties>().First();
                    cellProperties.Shading = new Shading() { Fill = "FFFF00" };
                }
                else
                {
                    cell.Append(new TableCellProperties(new Shading() { Fill = "FFFF00" }));
                }
            }
        }

        private void RemoveNinthColumnFromTargetTable(Table table)
        {
            foreach (var row in table.Elements<TableRow>())
            {
                var cells = row.Elements<TableCell>().ToList();
                if (cells.Count >= 9)
                {
                    cells[_slaughterColumnIndex].Remove();
                }
            }
        }
    }
}
