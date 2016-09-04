using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data.SqlClient;
using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;
using System.Web;
using System.Drawing.Printing;
using System.Threading;

namespace Zebra
{
    public partial class Form1 : Form
    {

        private List<String> arrayOne = new List<string>();
        private List<String> arrayTwo = new List<string>();
        private List<String> arrayThree = new List<string>();

        private string addProductsExcelFile = "";

        OleDbCommand command;
        OleDbDataAdapter da;
        private BindingSource bindingSource = null;
        private OleDbCommandBuilder oleCommandBuilder = null;
        DataTable dataTable = new DataTable();

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern SafeFileHandle CreateFile(string lpFileName, FileAccess dwDesiredAccess,
        uint dwShareMode, IntPtr lpSecurityAttributes, FileMode dwCreationDisposition,
        uint dwFlagsAndAttributes, IntPtr hTemplateFile);

        public Form1()
        {
            InitializeComponent();
        }

        private string resultCommand;
        private int rowCount;

        private OleDbConnection con = new OleDbConnection();
        //OleDbDataAdapter adap;
        DataSet ds;
        //OleDbCommandBuilder commandBuilder;

        private void Form1_Load(object sender, EventArgs e)
        {
            //try
            //{
            //    con.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=../../db.mdb;Persist Security Info=False;";
            //    con.Open();
            //    da = new OleDbDataAdapter("select * from goods", con);
            //    ds = new DataSet();
            //    da.Fill(ds);
            //    dataGridView1.DataSource = ds.Tables[0];
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
            //dataGridView1.EndEdit(); //very important step
            //da.Update(dataTable);
            DataBind();
            dataGridView1.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(dataGridView1_EditingControlShowing);

        }

        private void DataBind()
        {
            dataGridView1.DataSource = null;
            dataTable.Clear();

            String connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=../../db.mdb;Persist Security Info=False;";
            String queryString1 = "SELECT * FROM goods";

            OleDbConnection connection = new OleDbConnection(connectionString);
            connection.Open();
            OleDbCommand command = connection.CreateCommand();
            command.CommandText = queryString1;
            try
            {
                da = new OleDbDataAdapter(queryString1, connection);
                oleCommandBuilder = new OleDbCommandBuilder(da);
                da.Fill(dataTable);
                bindingSource = new BindingSource { DataSource = dataTable };
                dataGridView1.DataSource = bindingSource;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Оберіть файл";
            fdlg.InitialDirectory = @"c:\";
            fdlg.Filter = "Excel files (*.xls*)|*.xls*|Excel files (*.xls*)|*.xls*";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                addProductsExcelFile = fdlg.FileName;
            }
            Console.WriteLine(addProductsExcelFile);

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook excelBook = xlApp.Workbooks.Open(addProductsExcelFile);

            String[] excelSheets = new String[excelBook.Worksheets.Count];
            int i = 0;
            foreach (Microsoft.Office.Interop.Excel.Worksheet wSheet in excelBook.Worksheets)
            {
                excelSheets[i] = wSheet.Name;
                i++;
            }

            DataTable excelDataTable = ExcelToDataTable(addProductsExcelFile, excelSheets[0]);

            string[] columnNames = excelDataTable.Columns.Cast<DataColumn>()
                                 .Select(x => x.ColumnName)
                                 .ToArray();

            WorkWithDatabase workWithDatabase = new WorkWithDatabase();

            if (workWithDatabase.isConnection())
            {
                foreach (DataRow row in excelDataTable.Rows)
                {
                    //назва, товару, розмір етикетки, виробник, склад, додаткова інформація
                    string name = row[columnNames[0]].ToString();
                    Console.Write(name);

                    string size = row[columnNames[1]].ToString();
                    Console.Write(" " + size);

                    string manufacturer = row[columnNames[2]].ToString();
                    Console.Write(" " + manufacturer);

                    string consistance = row[columnNames[3]].ToString();
                    Console.Write(" " + consistance);

                    string additional = row[columnNames[4]].ToString();
                    Console.WriteLine(" " + additional);

                    workWithDatabase.insertData(name, size, manufacturer, consistance, additional);
                }
            }
            else
            {
                MessageBox.Show("Disconnected from database");
            }

            //this.Hide();

            //Form1 form = new Form1();
            //form.ShowDialog();
            //this.Close();

            DataTable dt = workWithDatabase.fillDataTable();

            BindingSource bs = goodsBindingSource;

            bs.DataSource = dt;

            dataGridView1.DataSource = bs;

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Оберіть файл";
            fdlg.InitialDirectory = @"c:\";
            fdlg.Filter = "Excel files (*.xls*)|*.xls*|Excel files (*.xls*)|*.xls*";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                addProductsExcelFile = fdlg.FileName;
            }
            Console.WriteLine(addProductsExcelFile);

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook excelBook = xlApp.Workbooks.Open(addProductsExcelFile);

            String[] excelSheets = new String[excelBook.Worksheets.Count];
            int i = 0;
            foreach (Microsoft.Office.Interop.Excel.Worksheet wSheet in excelBook.Worksheets)
            {
                excelSheets[i] = wSheet.Name;
                i++;
            }

            DataTable excelDataTable = ExcelToDataTable(addProductsExcelFile, excelSheets[0]);

            string[] columnNames = excelDataTable.Columns.Cast<DataColumn>()
                                 .Select(x => x.ColumnName)
                                 .ToArray();

            WorkWithDatabase workWithDatabase = new WorkWithDatabase();

            if (workWithDatabase.isConnection())
            {
                foreach (DataRow row in excelDataTable.Rows)
                {
                    string code = row[columnNames[0]].ToString();

                    string name = row[columnNames[1]].ToString();

                    string amount = row[columnNames[2]].ToString();

                    string size = "";
                    string manufacturer = "";
                    string composition = "";
                    string additional = "";

                    for (int rows = 0; rows < dataGridView1.Rows.Count - 1; rows++)
                    {
                        if (dataGridView1.Rows[rows].Cells[1].Value.ToString().Equals(name))
                        {
                            //for (int col = 0; col < dataGridView1.Rows[rows].Cells.Count; col++)
                            //{
                            size = dataGridView1.Rows[rows].Cells[2].Value.ToString();

                            manufacturer = dataGridView1.Rows[rows].Cells[3].Value.ToString();

                            composition = dataGridView1.Rows[rows].Cells[4].Value.ToString();

                            additional = dataGridView1.Rows[rows].Cells[5].Value.ToString();
                            //}
                        }
                    }

                    printMark(name, size, manufacturer, composition, additional, amount, code);
                }
            }
            else
            {
                MessageBox.Show("Disconnected from database");
            }

            if (MessageBox.Show("Починаю друкувати. Вставте етикетки розміром 3х5", "Друк", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                foreach (string stringToPrint in arrayOne)
                {
                    Print(stringToPrint);
                }
            }

            Thread.Sleep(2000);

            if (MessageBox.Show("Починаю друкувати. Вставте етикетки розміром 4х5", "Друк", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                foreach (string stringToPrint in arrayTwo)
                {
                    Print(stringToPrint);
                }
            }

            Thread.Sleep(2000);

            if (MessageBox.Show("Починаю друкувати. Вставте етикетки розміром 5х5", "Друк", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                foreach (string stringToPrint in arrayThree)
                {
                    Print(stringToPrint);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.EndEdit(); //very important step
            da.Update(dataTable);
            MessageBox.Show("Updated");
            DataBind();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                string value1 = row.Cells[0].Value.ToString();
                string value2 = row.Cells[1].Value.ToString();
                string value3 = row.Cells[2].Value.ToString();
                string value4 = row.Cells[3].Value.ToString();
                string value5 = row.Cells[4].Value.ToString();
                string value6 = row.Cells[5].Value.ToString();

                string value7 = "";
                if (GetAmount.InputBox("Кількість етикеток " + value2, "Кількість:", ref value7) != DialogResult.OK)
                    return;
                printMark(value2, value3, value4, value5, value6, value7, "");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string searchValue = "";
            if (GetAmount.InputBox("Пошук", "Введіть ключове слово або фразу:", ref searchValue) != DialogResult.OK)
                return;
            //int rowIndex = -1;

            //DataGridViewRow row = dataGridView1.Rows.Cast<DataGridViewRow>().Where(r => r.Cells["Names"].Value.ToString().Equals(searchValue)).First();
            dataGridView1.ClearSelection();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if ((row.Cells[1].Value != null) && (row.Cells[1].Value.ToString().Equals(searchValue)))
                    dataGridView1.Rows[row.Index].Selected = true;
                if ((row.Cells[1].Value != null) && (row.Cells[2].Value.ToString().Equals(searchValue)))
                    dataGridView1.Rows[row.Index].Selected = true;
                if ((row.Cells[1].Value != null) && (row.Cells[3].Value.ToString().Equals(searchValue)))
                    dataGridView1.Rows[row.Index].Selected = true;
                if ((row.Cells[1].Value != null) && (row.Cells[4].Value.ToString().Equals(searchValue)))
                    dataGridView1.Rows[row.Index].Selected = true;
                if ((row.Cells[1].Value != null) && (row.Cells[5].Value.ToString().Equals(searchValue)))
                    dataGridView1.Rows[row.Index].Selected = true;
            }
        }

        public DataTable ExcelToDataTable(string pathName, string sheetName)
        {
            DataTable tbContainer = new DataTable();
            string strConn = string.Empty;
            if (string.IsNullOrEmpty(sheetName)) { sheetName = "Sheet1"; }
            FileInfo file = new FileInfo(pathName);
            if (!file.Exists) { throw new Exception("Error, file doesn't exists!"); }
            string extension = file.Extension;
            switch (extension)
            {
                case ".xls":
                    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                    break;
                case ".xlsx":
                    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'";
                    break;
                default:
                    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                    break;
            }
            OleDbConnection cnnxls = new OleDbConnection(strConn);
            OleDbDataAdapter oda = new OleDbDataAdapter(string.Format("select * from [{0}$]", sheetName), cnnxls);
            DataSet ds = new DataSet();
            oda.Fill(tbContainer);
            return tbContainer;
        }

        public string checkForEnter(string stringToEdit)
        {
            string resultString = stringToEdit;
            if (stringToEdit.Contains("\r\n"))
            {
                resultString = "";
                string[] splittedStrings = stringToEdit.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                for (int i = 0; i < splittedStrings.Length; i++)
                {
                    while (splittedStrings[i].Length <= 63 && splittedStrings[i].Length >= 1)
                        splittedStrings[i] += " ";
                    resultString += splittedStrings[i];
                }
            }
            return resultString;
        }

        public void printMark(string name, string size, string manufacturer, string cont, string additional, string amount, string code)
        {
            resultCommand = "";
            if (additional.Equals(""))
            {
                Console.WriteLine("I'll not print this");
                return;
            }

            //checking what to print

            DialogResult dialogResult = MessageBox.Show("Друкую - " + name + " " + size + " " + manufacturer + " " + cont + " " + additional + " . " + amount + " штук.", "Друк", MessageBoxButtons.YesNo);
            if (dialogResult != DialogResult.Yes)
                return;

            Console.WriteLine("PRINTING - " + name + " " + size + " " + manufacturer + " " + cont + " " + additional + " " + amount + " " + code);
            resultCommand += "CT~~CD,~CC^~CT~^XA~TA000~JSN^LT0^MNW^MTD^PON^PMN^LH0,0^JMA^PR5,5~SD15^JUS^LRN^CI0^XZ^XA^MMT^PW400^LL0320^LS0^CWT,E:TT0003M_.FNT^CI28";
            rowCount = 0;

            addStringToPrinter(name);
            addStringToPrinter(manufacturer);
            addStringToPrinter(cont);
            addStringToPrinter(additional);

            resultCommand += "^PQ1,0,1,Y^XZ";
            Console.WriteLine(resultCommand);

            int amountMarks = Int32.Parse(amount);
            if (size.Equals("3:5"))
            {
                for (int i = 0; i < amountMarks; i++)
                {
                    arrayOne.Add(resultCommand);
                }
            }
            else if (size.Equals("4:5"))
            {
                for (int i = 0; i < amountMarks; i++)
                {
                    arrayTwo.Add(resultCommand);
                }
            }
            else if (size.Equals("5:5"))
            {
                for (int i = 0; i < amountMarks; i++)
                {
                    arrayThree.Add(resultCommand);
                }
            }

            //UNCOMMENT!!!!!!!!!!!!!!!!!!!!!

            //Print(resultCommand);
            MessageBox.Show(resultCommand);
        }

        private void Print(string s)
        {            
            s = changeCyrillic(s);
            PrintDialog pd = new PrintDialog();
            pd.PrinterSettings = new PrinterSettings();
            Printing.SendStringToPrinter(pd.PrinterSettings.PrinterName, s);
        }

        private void addStringToPrinter(string s)
        {
            s = checkForEnter(s);
            while (s.Length > 63)
            {
                double count = s.Length / 63;
                double truncated = Math.Truncate(count);
                for (double j = 0; j < truncated; j++)
                {
                    string tmp = s;
                    resultCommand += "^FT3," + (27 + (15 * rowCount)) + "^A0N,15,14^FH^FD" + tmp.Substring(0, 63) + "^FS";
                    rowCount++;
                    s = s.Substring(63);
                }
            }
            resultCommand += "^FT3," + (27 + (15 * rowCount)) + "^A0N,15,14^FH^FD" + s + "^FS";
            Console.WriteLine(s);
            Console.WriteLine(resultCommand);
            rowCount++;
        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
                TextBox txb = e.Control as TextBox;
                txb.PreviewKeyDown += (S, E) => {
                    if (E.KeyCode == Keys.Enter) {
                        txb.Text += Environment.NewLine;
                    }
                };
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;  
            }
        }

        //==============================================================================

        private string changeCyrillic(string s)
        {
            s = s.Replace("А", "_D0_90");
            s = s.Replace("Б", "_D0_91");
            s = s.Replace("В", "_D0_92");
            s = s.Replace("Г", "_D0_93");
            s = s.Replace("Ґ", "_D0_89");
            s = s.Replace("Д", "_D0_94");
            s = s.Replace("Е", "_D0_95");
            s = s.Replace("Є", "_D0_84");
            s = s.Replace("Ж", "_D0_96");
            s = s.Replace("З", "_D0_97");
            s = s.Replace("И", "_D0_98");
            s = s.Replace("І", "_D0_86");
            s = s.Replace("Ї", "_D0_87");
            s = s.Replace("Й", "_D0_99");
            s = s.Replace("К", "_D0_9A");
            s = s.Replace("Л", "_D0_9B");
            s = s.Replace("М", "_D0_9C");
            s = s.Replace("Н", "_D0_9D");
            s = s.Replace("О", "_D0_9E");
            s = s.Replace("П", "_D0_9F");
            s = s.Replace("Р", "_D0_A0");
            s = s.Replace("С", "_D0_A1");
            s = s.Replace("Т", "_D0_A2");
            s = s.Replace("У", "_D0_A3");
            s = s.Replace("Ф", "_D0_A4");
            s = s.Replace("Х", "_D0_A5");
            s = s.Replace("Ц", "_D0_A6");
            s = s.Replace("Ч", "_D0_A7");
            s = s.Replace("Ш", "_D0_A8");
            s = s.Replace("Щ", "_D0_A9");
            s = s.Replace("ь", "_D0_AC");
            s = s.Replace("Ю", "_D0_AE");
            s = s.Replace("Я", "_D0_AF");

            s = s.Replace("а", "_D0_B0");
            s = s.Replace("б", "_D0_B1");
            s = s.Replace("в", "_D0_B2");
            s = s.Replace("г", "_D0_B3");
            s = s.Replace("ґ", "_D2_91");
            s = s.Replace("д", "_D0_B4");
            s = s.Replace("е", "_D0_B5");
            s = s.Replace("є", "_D1_94");
            s = s.Replace("ж", "_D0_B6");
            s = s.Replace("з", "_D0_B7");
            s = s.Replace("и", "_D0_B8");
            s = s.Replace("і", "_D1_96");
            s = s.Replace("ї", "_D1_97");
            s = s.Replace("й", "_D0_B9");
            s = s.Replace("к", "_D0_BA");
            s = s.Replace("л", "_D0_BB");
            s = s.Replace("м", "_D0_BC");
            s = s.Replace("н", "_D0_BD");
            s = s.Replace("о", "_D0_BE");
            s = s.Replace("п", "_D0_BF");
            s = s.Replace("р", "_D1_80");
            s = s.Replace("с", "_D1_81");
            s = s.Replace("т", "_D1_82");
            s = s.Replace("у", "_D1_83");
            s = s.Replace("ф", "_D1_84");
            s = s.Replace("х", "_D1_85");
            s = s.Replace("ц", "_D1_86");
            s = s.Replace("ч", "_D1_87");
            s = s.Replace("ш", "_D1_88");
            s = s.Replace("щ", "_D1_89");
            s = s.Replace("ю", "_D1_8E");
            s = s.Replace("я", "_D1_8F");

            return s;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            Environment.Exit(1);
        }

        //private void button1_Click(object sender, EventArgs e)
        //{
        //    OpenFileDialog fdlg = new OpenFileDialog();
        //    fdlg.Title = "Оберіть файл";
        //    fdlg.InitialDirectory = @"c:\";
        //    fdlg.Filter = "Excel files (*.xls*)|*.xls*|Excel files (*.xls*)|*.xls*";
        //    fdlg.FilterIndex = 2;
        //    fdlg.RestoreDirectory = true;
        //    if (fdlg.ShowDialog() == DialogResult.OK)
        //    {
        //        addProductsExcelFile = fdlg.FileName;
        //    }
        //    Console.WriteLine(addProductsExcelFile);

        //    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        //    Microsoft.Office.Interop.Excel.Workbook excelBook = xlApp.Workbooks.Open(addProductsExcelFile);

        //    String[] excelSheets = new String[excelBook.Worksheets.Count];
        //    int i = 0;
        //    foreach (Microsoft.Office.Interop.Excel.Worksheet wSheet in excelBook.Worksheets)
        //    {
        //        excelSheets[i] = wSheet.Name;
        //        i++;
        //    }

        //    DataTable excelDataTable = ExcelToDataTable(addProductsExcelFile, excelSheets[0]);

        //    string[] columnNames = excelDataTable.Columns.Cast<DataColumn>()
        //                         .Select(x => x.ColumnName)
        //                         .ToArray();

        //    WorkWithDatabase workWithDatabase = new WorkWithDatabase();

        //    if (workWithDatabase.isConnection())
        //    {
        //        foreach (DataRow row in excelDataTable.Rows)
        //        {
        //            //назва, товару, розмір етикетки, виробник, склад, додаткова інформація
        //            string name = row[columnNames[0]].ToString();
        //            Console.Write(name);

        //            string size = row[columnNames[1]].ToString();
        //            Console.Write(" " + size);

        //            string manufacturer = row[columnNames[2]].ToString();
        //            Console.Write(" " + manufacturer);

        //            string consistance = row[columnNames[3]].ToString();
        //            Console.Write(" " + consistance);

        //            string additional = row[columnNames[4]].ToString();
        //            Console.WriteLine(" " + additional);

        //            workWithDatabase.insertData(name, size, manufacturer, consistance, additional);
        //        }
        //    }
        //    else
        //    {
        //        MessageBox.Show("Disconnected from database");
        //    }

        //    //this.Hide();

        //    //Form1 form = new Form1();
        //    //form.ShowDialog();
        //    //this.Close();

        //    DataTable dt = workWithDatabase.fillDataTable();

        //    BindingSource bs = goodsBindingSource;

        //    bs.DataSource = dt;

        //    dataGridView1.DataSource = bs;

        //}

        //private void button2_Click(object sender, EventArgs e)
        //{
        //    OpenFileDialog fdlg = new OpenFileDialog();
        //    fdlg.Title = "Оберіть файл";
        //    fdlg.InitialDirectory = @"c:\";
        //    fdlg.Filter = "Excel files (*.xls*)|*.xls*|Excel files (*.xls*)|*.xls*";
        //    fdlg.FilterIndex = 2;
        //    fdlg.RestoreDirectory = true;
        //    if (fdlg.ShowDialog() == DialogResult.OK)
        //    {
        //        addProductsExcelFile = fdlg.FileName;
        //    }
        //    Console.WriteLine(addProductsExcelFile);

        //    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        //    Microsoft.Office.Interop.Excel.Workbook excelBook = xlApp.Workbooks.Open(addProductsExcelFile);

        //    String[] excelSheets = new String[excelBook.Worksheets.Count];
        //    int i = 0;
        //    foreach (Microsoft.Office.Interop.Excel.Worksheet wSheet in excelBook.Worksheets)
        //    {
        //        excelSheets[i] = wSheet.Name;
        //        i++;
        //    }

        //    DataTable excelDataTable = ExcelToDataTable(addProductsExcelFile, excelSheets[0]);

        //    string[] columnNames = excelDataTable.Columns.Cast<DataColumn>()
        //                         .Select(x => x.ColumnName)
        //                         .ToArray();

        //    WorkWithDatabase workWithDatabase = new WorkWithDatabase();

        //    if (workWithDatabase.isConnection())
        //    {
        //        foreach (DataRow row in excelDataTable.Rows)
        //        {
        //            string code = row[columnNames[0]].ToString();

        //            string name = row[columnNames[1]].ToString();

        //            string amount = row[columnNames[2]].ToString();

        //            string size = "";
        //            string manufacturer = "";
        //            string composition = "";
        //            string additional = "";

        //            for (int rows = 0; rows < dataGridView1.Rows.Count - 1; rows++)
        //            {
        //                if (dataGridView1.Rows[rows].Cells[1].Value.ToString().Equals(name))
        //                {
        //                    //for (int col = 0; col < dataGridView1.Rows[rows].Cells.Count; col++)
        //                    //{
        //                    size = dataGridView1.Rows[rows].Cells[2].Value.ToString();

        //                    manufacturer = dataGridView1.Rows[rows].Cells[3].Value.ToString();

        //                    composition = dataGridView1.Rows[rows].Cells[4].Value.ToString();

        //                    additional = dataGridView1.Rows[rows].Cells[5].Value.ToString();
        //                    //}
        //                }
        //            }

        //            printMark(name, size, manufacturer, composition, additional, amount, code);
        //        }
        //    }
        //    else
        //    {
        //        MessageBox.Show("Disconnected from database");
        //    }

        //    if (MessageBox.Show("Починаю друкувати. Вставте етикетки розміром 3х5", "Друк", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
        //    {
        //        foreach (string stringToPrint in arrayOne)
        //        {
        //            Print(stringToPrint);
        //        }
        //    }

        //    Thread.Sleep(2000);

        //    if (MessageBox.Show("Починаю друкувати. Вставте етикетки розміром 4х5", "Друк", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
        //    {
        //        foreach (string stringToPrint in arrayTwo)
        //        {
        //            Print(stringToPrint);
        //        }
        //    }

        //    Thread.Sleep(2000);

        //    if (MessageBox.Show("Починаю друкувати. Вставте етикетки розміром 5х5", "Друк", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
        //    {
        //        foreach (string stringToPrint in arrayThree)
        //        {
        //            Print(stringToPrint);
        //        }
        //    }

        //    //Print();
        //}

    }
}
