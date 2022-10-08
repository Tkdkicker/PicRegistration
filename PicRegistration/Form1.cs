using Microsoft.VisualBasic;
using System;
using System.Data;
using System.Data.Odbc;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PicRegistration
{
    public partial class Form1 : Form
    {
        #region Private fields

        private bool _closing = false;
        private DataTable _table = new DataTable();
        private string _wafer = string.Empty;

        #endregion

        #region Constructor

        public Form1()
        {
            InitializeComponent();
        }

        #endregion

        #region Form1_Load

        private void Form1_Load(object sender, EventArgs e)
        {
            // Define data table.
            _table = new DataTable();
            _table.Columns.Add("OSA", typeof(string));
            _table.Columns.Add("Chip", typeof(string));
            _table.Columns.Add("GRN", typeof(string));

            // Set data source.
            dataGridView1.DataSource = _table;
        }

        #endregion

        #region dataGridView1_CellValidating

        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (_closing)
                return;

            // Special case for GRN numbers.
            if (e.ColumnIndex == 2)
            {
                // Set all rows below with the entered value.
                for (int row = e.RowIndex + 1; row < _table.Rows.Count; row++)
                {
                    _table.Rows[row][e.ColumnIndex] = e.FormattedValue;
                }

                return;
            }

            // Enumerate all rows.
            for (int row = 0; row < _table.Rows.Count; row++)
            {
                // Check row is different and if contents are the same.
                if (row != e.RowIndex && string.Equals(_table.Rows[row][e.ColumnIndex], e.FormattedValue))
                {
                    // Show message box and cancel.
                    MessageBox.Show("Item " + e.FormattedValue + " already exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                    break;
                }
            }
        }

        #endregion

        #region button1_Click

        private void button1_Click(object sender, EventArgs e)
        {
            // Check for any data.
            if (_table.Rows.Count == 0)
            {
                MessageBox.Show("Please enter PIC registration data", "Missing data", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Define abort.
            bool abort = false;

            // Enumerate all rows and columns.
            for (int row = 0; row < _table.Rows.Count; row++)
            {
                // Change to enumerate first two columns (ignoring GRN column).
                for (int column = 0; column < 2; column++)
                {
                    // Determine if cell is empty.
                    bool empty = string.IsNullOrEmpty(_table.Rows[row].Field<string>(column));

                    // Select empty cells.
                    dataGridView1.Rows[row].Cells[column].Selected = empty;

                    if (empty)
                        abort = true;
                }
            }

            if (abort)
            {
                MessageBox.Show("Please enter missing data", "Missing data", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Define temporary file.
            string file = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory, Environment.SpecialFolderOption.None), "upload.csv");

            if (File.Exists(file))
            {
                try
                {
                    File.Delete(file);
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // Save as CSV.
            SaveAsCSV(_table, file);

            // Get shop order.
            DataRow shop = GetShopOrder(_table.Rows[0].Field<string>(0));

            if (shop == null)
            {
                // Save to archive.
                SaveToArchive(file);
            }
            else
            {
                // Open web page.
                OpenWebPage(shop.Field<int>("id"));

                // Save to archive.
                SaveToArchive(file, shop.Field<string>("name"));
            }

            // Show message box.
            MessageBox.Show("File has been saved to desktop", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion

        #region SaveToArchive

        private static void SaveToArchive(string source)
        {
            SaveToArchive(source, null);
        }

        private static void SaveToArchive(string source, string name)
        {
            // Check for valid parameters.
            if (string.IsNullOrEmpty(source))
                throw new ArgumentException("Source file cannot be null or empty", nameof(source));

            if (string.IsNullOrEmpty(name))
                name = "UNKNOWN";

            // Define path.
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory, Environment.SpecialFolderOption.None), "PIC registration");

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            // Define file.
            string file = Path.Combine(path, name + " " + DateTime.Now.ToString("yyyyMMddhhmmss") + Path.GetExtension(source));

            // Copy file.
            File.Copy(source, file);
        }

        #endregion

        #region OpenWebPage

        private static Process OpenWebPage(int id)
        {
            // Open web page.
            return Process.Start("http://wiptracker.ep.lan/stations/bulk_pic_registration/" + id + "/");
        }

        #endregion

        #region GetShopOrder

        private static DataRow GetShopOrder(string serial)
        {
            // Open ODBC connection to database.
            using (OdbcConnection connection = new OdbcConnection(GetConnectionString()))
            {
                try
                {
                    connection.Open();
                }
                catch
                {
                    return null;
                }

                // Define command to find shop order ID from first serial number specified.
                using (OdbcCommand command = new OdbcCommand("SELECT d.id, d.name FROM public.osa a LEFT JOIN public.device b ON b.osa_id = a.id LEFT JOIN public.tracker c ON c.device_id = b.id LEFT JOIN public.shop_order d ON d.id = c.shoporder_id WHERE a.name = ? ", connection))
                {
                    // Add criteria.
                    command.Parameters.AddWithValue("?", serial);

                    // Define data table.
                    DataTable table = new DataTable();

                    // Define data adapter.
                    using (OdbcDataAdapter adapter = new OdbcDataAdapter(command))
                    {
                        adapter.Fill(table);

                        if (table.Rows.Count == 0)
                            return null;

                        // Return first row.
                        return table.Rows[0];
                    }
                }
            }
        }

        #endregion

        #region SaveAsCSV

        private void SaveAsCSV(DataTable table, string file)
        {
            // Check for valid parameters.
            if (table == null)
                throw new ArgumentNullException(nameof(table));

            if (string.IsNullOrEmpty(file))
                throw new ArgumentException("File cannot be null or empty", nameof(file));

            // Define string builder.
            StringBuilder text = new StringBuilder();

            // Append header.
            text.Append("OSA,CHIP,BATCH_NUMBER");

            // Enumerate rows.
            for (int row = 0; row < table.Rows.Count; row++)
            {
                // Append new line.
                text.Append(Environment.NewLine);

                // Enumerate columns.
                for (int column = 0; column < table.Columns.Count; column++)
                {
                    // Get cell.
                    string cell = table.Rows[row][column]?.ToString();
                    
                    // Special case for Chip column to append wafer number.
                    if (column == 1 && !cell.StartsWith("CD"))
                    {
                        // Check for trailing digit.
                        if (char.IsDigit(cell[cell.Length - 1]))
                            cell += "_A";
                        else
                            cell = cell.Substring(0, cell.Length - 1) + "_" + cell.Substring(cell.Length - 1);

                        // Add "CD" prefix and wafer suffix.
                        cell = "CD" + cell + "-" + _wafer;
                    }

                    // Enumerate rows.
                    if (column > 0)
                        text.Append(',');

                    // Add cell.
                    text.Append(cell);
                }
            }

            // Write contents.
            File.WriteAllText(file, text.ToString());
        }

        #endregion

        #region Form1_Shown

        private void Form1_Shown(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_wafer))
                return;

            // Infinate loop while wafer does not match required regular expression.
            while (!Regex.IsMatch(_wafer, @"^(\d{5}-\d{3})|(\D{2}\d{6,})|([0-9A-Z\-]{6,})$"))
            {
                // Check if wafer has been specified.
                if (!string.IsNullOrEmpty(_wafer))
                    MessageBox.Show("Incorrect wafer number format", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Error);

                // Ask user for wafer number.
                _wafer = Interaction.InputBox("Please enter wafer number in the correct format", "Wafer", _wafer);

                // Check for null wafer.
                if (string.IsNullOrEmpty(_wafer))
                {
                    // Prompt user to enter wafer number.
                    if (MessageBox.Show("You need to enter the wafer number", "No input", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                    {
                        Close();
                        return;
                    }
                }
            }

            // Change title.
            Text += " " + _wafer;
        }

        #endregion

        #region Form1_FormClosing

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Check if closure due to user and there is data.
            if (e.CloseReason == CloseReason.UserClosing && _table.Rows.Count > 0 && !_closing)
                e.Cancel = MessageBox.Show("Are you sure you want to exit?", "Close application", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No;

            // Set closing flag.
            _closing = !e.Cancel;
        }

        #endregion

        #region GetConnectionString

        internal static string GetConnectionString()
        {
            #region Confidential
            #region Confidential
            #region Confidential
            return "Dsn=BI;Uid=djangoclient@effect-postgres;Pwd=WDM-Pon/postgres";
            #endregion
            #endregion
            #endregion
        }

        #endregion
    }
}
