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
         */

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

        private void button2_Click(object sender, EventArgs e)
        {
            var connection = new SQLiteConnection(@"Data Source=D:\Data\Download\sqlite\test.db");
//            var context = new DataContext(connection);
//
//            var companies = context.GetTable<Company>();
//            foreach (var company in companies)
//            {
//                Console.WriteLine("Company: {0} {1}",
//                    company.Id, company.Seats);
//            }
            connection.Open();
            SQLiteCommand liteCommand = new SQLiteCommand("select * from barang", connection);
            SQLiteDataReader reader = liteCommand.ExecuteReader();

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

    }
}
