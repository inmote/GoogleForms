using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Text = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace GoogleForms
{
    public partial class MainForm : Form
    {
        private String mCSVInputfile;
        private String mXLSTemplatefile;
        private String mXLSOutputFolder;
        private Dictionary<String, String> map = new Dictionary<String, String>
        {
            { "\"Zeer mee oneens\"",       "C" },
            { "\"Mee oneens\"",            "D" },
            { "\"Gedeeltelijk mee eens\"", "E" },
            { "\"Mee eens\"",              "F" },
            { "\"Zeer mee eens\"",         "G" },
            { "\"Geen antwoord\"",         "H" }
        };

        public MainForm()
        {
            InitializeComponent();
            Text = "Google Forms v" + Assembly.GetExecutingAssembly().GetName().Version.ToString();
            ReadSettings();
            if (File.Exists(mCSVInputfile))
            {
                tbCsvInputFile.Text = mCSVInputfile;
            }
            else
            {
                tbCsvInputFile.Text = mCSVInputfile = "";
            }
            if (File.Exists(mXLSTemplatefile))
            {
                tbXLSTemplate.Text = mXLSTemplatefile;
            }
            else
            {
                tbXLSTemplate.Text = mXLSTemplatefile = "";
            }
            if (Directory.Exists(mXLSOutputFolder))
            {
                tbXLSOutputFolder.Text = mXLSOutputFolder;
            }
            else
            {
                tbXLSOutputFolder.Text = mXLSOutputFolder = "";
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                InitialDirectory = @"C:\",
                Title = "Select CSV input files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "csv",
                Filter = "CSV files (*.csv)|*.csv",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                mCSVInputfile = tbCsvInputFile.Text = openFileDialog.FileName;
                WriteSettings();
            }
        }
        Cell GetSpreadsheetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>()?.Elements<Row>().Where(r => r.RowIndex == rowIndex);
            if (rows is null || rows.Count() == 0)
            {
                // A cell does not exist at the specified row.
                return null;
            }
            IEnumerable<Cell> cells = rows.First().Elements<Cell>().Where(c => string.Compare(c.CellReference?.Value, columnName + rowIndex, true) == 0);
            if (cells.Count() == 0)
            {
                // A cell does not exist at the specified column, in the specified row.
                return null;
            }
            return cells.FirstOrDefault();
        }

        private Cell CreateCell(String cellReference)
        {
            Cell cell = new Cell
            {
                CellReference = cellReference,
                DataType = CellValues.InlineString
            };

            InlineString inlineString = new InlineString();
            Text t = new Text
            {
                Text = "X"
            };
            inlineString.AppendChild(t);

            cell.AppendChild(inlineString);

            return cell;
        }

        private Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            String cellReference = columnName + rowIndex;

            Row row;
            // See if the row already exists.
            if (sheetData?.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // Find the cell.
            Cell newCell = null;
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                newCell = row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();

                newCell.DataType = CellValues.InlineString;
                InlineString inlineString = new InlineString();
                Text t = new Text
                {
                    Text = "X"
                };
                inlineString.AppendChild(t);
                newCell.AppendChild(inlineString);
            }
            else
            {
                // Cells must be in sequential order according to CellReference.
                // Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell refCellCandidate in row.Elements<Cell>())
                {
                    if (string.Compare(refCellCandidate.CellReference?.Value, cellReference, true) > 0)
                    {
                        refCell = refCellCandidate;
                        break;
                    }
                }

                newCell = CreateCell(cellReference);
                row.InsertBefore(newCell, refCell);
            }

            return newCell;
        }

        private String DetermineColumnName(String answer)
        {
            String columnName = "";
            if (map.ContainsKey(answer))
            {
                columnName = map[answer];
            }

            return columnName;
        }

        private bool ValidateInputParams()
        {
            if (!File.Exists(mCSVInputfile))
            {
                MessageBox.Show("First set a valid CSV input file!", "Error notification");
                return false;
            }
            if (!File.Exists(mXLSTemplatefile))
            {
                MessageBox.Show("First set a valid XLSX template file!", "Error notification");
                return false;
            }
            if (!Directory.Exists(mXLSOutputFolder))
            {
                MessageBox.Show("First set a valid XLSX output folder!", "Error notification");
                return false;
            }
            return true;
        }
        private void btnGenerate_Click(object sender, EventArgs e)
        {
            if(!ValidateInputParams())
            {
                return;
            }

            int count = 0;
            Directory.Delete(mXLSOutputFolder, true);
            Directory.CreateDirectory(mXLSOutputFolder);
            using (StreamReader sr = new StreamReader(mCSVInputfile))
            {
                String line;
                pbProgress.Value = 0;
                // currentLine will be null when the StreamReader reaches the end of file
                while ((line = sr.ReadLine()) != null)
                {
                    if (count > 0)
                    {
                        String email = "";
                        List<String> answers = line.Split(',').ToList();
                        if (answers.Count == map.Count + 1)
                        {
                            email = answers[1].Replace('.', '_').Replace('@', '_');
                            email = email.Replace("\"", "");
                        }

                        String fileName = Path.GetFileName(mXLSTemplatefile);
                        String outputFileName = mXLSOutputFolder + "\\" + count;
                        if (!String.IsNullOrEmpty(email))
                        {
                            outputFileName += ("-" + email);
                        }
                        outputFileName += ("-" + fileName);
                        File.Copy(mXLSTemplatefile, outputFileName);
                        
                        using (SpreadsheetDocument doc = SpreadsheetDocument.Open(outputFileName, true))
                        {
                            WorkbookPart workbookPart = doc.WorkbookPart;
                            Sheet sheet = workbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                            WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));

                            uint rowNumber = 3;
                            String columnName = "";
                            foreach (String answer in answers)
                            {
                                columnName = DetermineColumnName(answer.Trim());
                                if (!String.IsNullOrEmpty(columnName))
                                {
                                    Cell cell = InsertCellInWorksheet(columnName, rowNumber, worksheetPart);
                                    rowNumber++;
                                }
                            }
                        }
                    }
                    count++;
                }
            }

            string message = (count - 1) + " XLS output files generated in " + mXLSOutputFolder + ".";
            string title = "Generate XLS output";
            MessageBox.Show(message, title);
        }

        private void btnOutputBrowse_Click(object sender, EventArgs e)
        {
                FolderBrowserDialog folderDlg = new FolderBrowserDialog
                {
                    ShowNewFolderButton = true
                };
                DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                mXLSOutputFolder = tbXLSOutputFolder.Text = folderDlg.SelectedPath;
                WriteSettings();
            }
        }

        private void ReadSettings()
        {
            mCSVInputfile    = ConfigurationManager.AppSettings["csv_input_file"];
            mXLSTemplatefile = ConfigurationManager.AppSettings["xls_template_file"];
            mXLSOutputFolder = ConfigurationManager.AppSettings["xls_output_folder"];
        }
        private void AddUpdateAppSettings(string key, string value)
        {
            try
            {
                var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var settings = configFile.AppSettings.Settings;
                if (settings[key] == null)
                {
                    settings.Add(key, value);
                }
                else
                {
                    settings[key].Value = value;
                }
                configFile.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
            }
            catch (ConfigurationErrorsException)
            {
                Console.WriteLine("Error writing app settings");
            }
        }
        private void WriteSettings()
        {
            AddUpdateAppSettings("csv_input_file",    mCSVInputfile);
            AddUpdateAppSettings("xls_template_file", mXLSTemplatefile);
            AddUpdateAppSettings("xls_output_folder", mXLSOutputFolder);
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            WriteSettings();
        }

        private void btnXLSBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                InitialDirectory = @"C:\",
                Title = "Select XLSX template file",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xlsx",
                Filter = "XLSX files (*.xlsx)|*.xlsx",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                mXLSTemplatefile = tbXLSTemplate.Text = openFileDialog.FileName;
                WriteSettings();
            }
        }
    }
}
