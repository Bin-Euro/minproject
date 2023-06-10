using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PRN211_Day4_ExcelHandling
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnChooseFolder_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog()
                == DialogResult.OK)
            {
                txtFolderName.Text = folderBrowserDialog1.SelectedPath;
                var xlsFiles = ScanFileInDirectory(txtFolderName.Text);
                foreach (var file in xlsFiles)
                {
                    listBox1.Items.Add(file);
                }
            }

        }

        /// <summary>
        /// Scan directory and return a list of file path with *.xls extension
        /// </summary>
        /// <param name="directoryPath"></param>
        /// <returns></returns>
        private List<string> ScanFileInDirectory(string directoryPath)
        {
            var files = Directory.GetFiles(directoryPath);
            var xlsFiles = new List<string>();
            foreach (var file in files)
            {
                if (file.EndsWith(".xls"))
                {
                    xlsFiles.Add(Path.GetFileName(file));
                }

            }
            return xlsFiles;
        }

        private void btnView_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex >= 0)
            {
                var filename = listBox1.SelectedItems[0];
                var fullPath = txtFolderName.Text + "\\" + filename;
                ParseExcel(fullPath);
            }
        }

        private List<StudentCourse> ParseExcel(string xlsFile)
        {
            //Open excel workbook
            //Get first sheet 
            // Scan each row
            //Read data and convert to StudentCourse object
            //Add to list 
            //Return 


            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(xlsFile, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                string text;
                foreach (Row r in sheetData.Elements<Row>())
                {
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        text = c.CellValue.Text;
                        
                    }
                }
            }
            return null;

        }
    }
}