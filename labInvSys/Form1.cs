using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Data.SQLite;

/* For I/O purpose */
using System.IO;

/* For Diagnostics */
using System.Diagnostics;

/* Imported custom DLL */
using _conectorSQLite;
using _encryption;

namespace labInvSys
{
    public partial class Form1 : Form
    {
        /*
         *  Proyek Manajemen Stok Barang Lab. Terpadu Poltekkes Kemenkes Surabaya
         *  Start : 18 Mei 2015 12:52
         *  Real Start : 18 Juni 2015 12:53
         *  Om Awignamastu Namah Siddham,
         *  
         *  Aplikasi Menejemen Stok Barang Lab. Terpadu Poltekkes Surabaya versi 0.1.1
         *  1. Digunakan untuk memantau, menginventaris barang habis pakai milik lab terpadu;
         *  2. Client mengolah data stok dengan menggunakan MS. Excel;
         *  3. Master membaca data dari client yg berupa file *.xlsx kemudian memprosesnya sebagai sebuah laporan.
         *  
         *  14/07/2015
         *  1. Selesai membuat .dll koneksi untuk SQLite
         *  
         */

        #region KOMPONEN WAJIB
        readonly CConectionSQLite _connect = new CConectionSQLite();
        private SQLiteConnection _connection;
        private string _sqlQuery;
        private readonly string _configurationManager = Properties.Settings.Default.Setting;
        #endregion

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var existingFile = new FileInfo(@"D:\new\Untitled 1.xlsx");
                // Open and read the XlSX file.
            using (var package = new ExcelPackage(existingFile))
            {
                // Get the work book in the file
                ExcelWorkbook workBook = package.Workbook;
                if (workBook != null)
                {
                    if (workBook.Worksheets.Count > 0)
                    {
                        // Get the first worksheet
                        ExcelWorksheet currentWorksheet = workBook.Worksheets.First();

                        // read some data
                        textBox1.Text = currentWorksheet.Cells[1, 1].Text;
                    }
                }
            }
        }

        private void selectData()
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Test");
//            var connection = new SQLiteConnection(@"Data Source=D:\new\sqlite\test.db");
//            var context = new DataContext(connection);
//
//            var companies = context.GetTable<Company>();
//            foreach (var company in companies)
//            {
//                Console.WriteLine("Company: {0} {1}",
//                    company.Id, company.Seats);
//            }
            string errMsg = "";

            _connection = _connect.Connect(_configurationManager, ref errMsg, "123");
            if (!string.IsNullOrEmpty(errMsg))
            {
                MessageBox.Show(errMsg);
                return;
            }
 

            SQLiteDataReader reader = _connect.ReadingSqLiteDataReader("select * from barang", _connection, ref errMsg);

            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ListViewItem item = new ListViewItem(reader[0].ToString());
                    item.SubItems.Add(reader[1].ToString());
                    item.SubItems.Add(reader[2].ToString());
                    listView1.Items.Add(item);
                }
                reader.Close();
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string errMsg = "";

            _connection = _connect.Connect(_configurationManager, ref errMsg, "123");
            if (!string.IsNullOrEmpty(errMsg))
            {
                MessageBox.Show(errMsg);
                return;
            }
            SQLiteTransaction sqLiteTransaction = _connection.BeginTransaction();

            var existingFile = new FileInfo(@"D:\new\Untitled 1.xlsx");
            // Open and read the XlSX file.
            try
            {
                using (var package = new ExcelPackage(existingFile))
                {
                    // Get the work book in the file
                    ExcelWorkbook workBook = package.Workbook;
                    if (workBook != null)
                    {
                        if (workBook.Worksheets.Count > 0)
                        {
                            // Get the first worksheet
                            ExcelWorksheet currentWorksheet = workBook.Worksheets.First();
                            var rowsCount = currentWorksheet.Dimension.End.Row;
                            for (int row = 2; row <= rowsCount; row++)
                            {
                                //                            MessageBox.Show(currentWorksheet.Cells[row, 1].Text + currentWorksheet.Cells[row, 2].Text +
                                //                                            currentWorksheet.Cells[row, 3].Text +
                                //                                            currentWorksheet.Cells[row, 4].Text + currentWorksheet.Cells[row, 5].Text +
                                //                                            currentWorksheet.Cells[row, 6].Text +
                                //                                            currentWorksheet.Cells[row, 7].Text + currentWorksheet.Cells[row, 8].Text +
                                //                                            currentWorksheet.Cells[row, 9].Text);
                                _connect.Insertion(
                                    "insert into barang values ('" + currentWorksheet.Cells[row, 1].Text +
                                    "', '" + currentWorksheet.Cells[row, 2].Text +
                                    "', '" + currentWorksheet.Cells[row, 3].Text +
                                    "', '" + currentWorksheet.Cells[row, 4].Text +
                                    "', '" + currentWorksheet.Cells[row, 5].Text +
                                    "', '" + currentWorksheet.Cells[row, 6].Text +
                                    "', '" + currentWorksheet.Cells[row, 7].Text +
                                    "', '" + currentWorksheet.Cells[row, 8].Text +
                                    "', '" + currentWorksheet.Cells[row, 9].Text + "')",
                                    _connection, sqLiteTransaction, ref errMsg);

                                if (!string.IsNullOrEmpty(errMsg))
                                {
                                    MessageBox.Show(errMsg);
                                    sqLiteTransaction.Rollback();
                                    _connection.Close();
                                    return;
                                }

                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"KESALAHAN", MessageBoxButtons.OK);
                return;
            }

            sqLiteTransaction.Commit();
            _connection.Close();
        }

    }
}
