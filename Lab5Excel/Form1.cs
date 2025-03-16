using OfficeOpenXml;

namespace Lab5Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filePath = @"C:\Users\Computer\Desktop\proba.xlsx";

            FileInfo file = new FileInfo(filePath);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault();

                if (worksheet == null)
                {
                    worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
                }

                worksheet.Cells[1, 1].Value = textBox1.Text;  // Scrie valoarea în celula A1
                excelPackage.Save();
            }
            richTextBox1.AppendText("Date scrise în Excel cu succes!\n");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string filePath = @"C:\Users\Computer\Desktop\proba.xlsx";
            FileInfo file = new FileInfo(filePath);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Setare licență EPPlus

            if (!file.Exists)
            {
                MessageBox.Show("Fișierul Excel nu există!");
                return;
            }

            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault();

                if (worksheet == null)
                {
                    MessageBox.Show("Nu există foi de calcul în fișierul Excel!");
                    return;
                }

                object readValue = worksheet.Cells[1, 1].Value;  // Citește valoarea din celula A1
                textBox2.Text = readValue != null ? readValue.ToString() : "Celula este goală";

                richTextBox1.AppendText("Date citite din Excel cu succes!\n");

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string filePath = @"C:\Users\Computer\Desktop\data.txt";

            if (!File.Exists(filePath))
            {
                MessageBox.Show("Fișierul .txt nu există!");
                return;
            }

            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            dataGridView1.Columns.Add("ProdID", "ProdID");
            dataGridView1.Columns.Add("Date1", "Date1");
            dataGridView1.Columns.Add("QuantS", "QuantS");

            string[] lines = File.ReadAllLines(filePath);

            foreach (string line in lines)
            {
                string[] values = line.Split(',');

                if (values.Length == 3) // Asigură că sunt 3 coloane
                {
                    dataGridView1.Rows.Add(values[0], values[1], values[2]);
                }
            }

            richTextBox1.AppendText("DataGridView a fost completat cu datele din txt file\n");

        }

        private void button4_Click(object sender, EventArgs e)
        {
            string filePath = @"C:\Users\Computer\Desktop\proba.xlsx";
            FileInfo file = new FileInfo(filePath);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Setare licență EPPlus

            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault();

                if (worksheet == null)
                {
                    worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
                }

                // Scriere antete în Excel
                for (int col = 0; col < dataGridView1.Columns.Count; col++)
                {
                    worksheet.Cells[1, col + 1].Value = dataGridView1.Columns[col].HeaderText;
                }

                // Scriere date în Excel
                for (int row = 0; row < dataGridView1.Rows.Count; row++)
                {
                    for (int col = 0; col < dataGridView1.Columns.Count; col++)
                    {
                        object value = dataGridView1.Rows[row].Cells[col].Value;
                        worksheet.Cells[row + 2, col + 1].Value = value != null ? value.ToString() : "";
                    }
                }

                excelPackage.Save();
            }

            richTextBox1.AppendText("Datele din DataGridView au fost salvate în Excel!\n");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string filePath = @"C:\Users\Computer\Desktop\test.txt";

            using (StreamWriter writer = new StreamWriter(filePath))
            {
                // Scrie anteturile coloanelor
                for (int col = 0; col < dataGridView2.Columns.Count; col++)
                {
                    writer.Write(dataGridView2.Columns[col].HeaderText);
                    if (col < dataGridView2.Columns.Count - 1)
                        writer.Write(","); // Separator CSV
                }
                writer.WriteLine();

                // Scrie datele rând cu rând
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (!row.IsNewRow) // Ignoră ultimul rând gol
                    {
                        for (int col = 0; col < dataGridView2.Columns.Count; col++)
                        {
                            writer.Write(row.Cells[col].Value?.ToString() ?? "");
                            if (col < dataGridView2.Columns.Count - 1)
                                writer.Write(",");
                        }
                        writer.WriteLine();
                    }
                }
            }

            richTextBox1.AppendText($"Datele din DataGridView au fost salvate în fișierul TXT! {filePath} \n");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string filePath = @"C:\Users\Computer\Desktop\proba.xlsx";

            FileInfo file = new FileInfo(filePath);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (!file.Exists)
            {
                MessageBox.Show("Fișierul Excel nu există!");
                return;
            }

            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault();

                if (worksheet == null)
                {
                    MessageBox.Show("Nu există foi de calcul în fișierul Excel!");
                    return;
                }

                dataGridView2.Rows.Clear();
                dataGridView2.Columns.Clear();

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                // Adaugă coloanele pe baza antetelor din Excel
                for (int col = 1; col <= colCount; col++)
                {
                    string header = worksheet.Cells[1, col].Value?.ToString() ?? $"Col{col}";
                    dataGridView2.Columns.Add($"Col{col}", header);
                }

                // Adaugă rândurile cu date
                for (int row = 2; row <= rowCount; row++)
                {
                    object[] rowData = new object[colCount];
                    for (int col = 1; col <= colCount; col++)
                    {
                        rowData[col - 1] = worksheet.Cells[row, col].Value?.ToString() ?? "";
                    }
                    dataGridView2.Rows.Add(rowData);
                }
            }

            richTextBox1.AppendText("Datele din Excel au fost încărcate în DataGridView!\n");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string filePath = @"C:\Users\Computer\Desktop\proba.xlsx";
            FileInfo file = new FileInfo(filePath);
            Random rnd = new Random();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault();

                if (worksheet == null)
                {
                    worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
                }

                // scriem valori random in coloana J (coloana 10) de la J1 la J12
                for (int row = 1; row <= 12; row++)
                {
                    worksheet.Cells[row, 10].Value = rnd.Next(10, 100);
                }

                // scriem formula de suma in celula data
                worksheet.Cells[1, 12].Formula = "SUM(J1:J12)";

                excelPackage.Save();

                worksheet.Calculate(); // calculam formula in excel
                object sumValue = worksheet.Cells[1, 12].Value;

                txtExcelSum.Text = sumValue != null ? sumValue.ToString() : "0";
            }

            richTextBox1.AppendText("Valori random înscrise în Excel și suma calculată! \n");
        }
    }
}
