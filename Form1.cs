using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace Costing
{
    public partial class Form1 : Form
    {
        private const string connectionString = @"Data Source=.\SQLEXPRESS;Initial Catalog=dbdemo;Integrated Security=True";
        private string selectedStyle;
        private string selectedCustomer;



        public Form1()
        {
            InitializeComponent();
            PopulateComboBox();
            // Call method to populate ComboBox when the form loads
            dgvColumns.CellMouseDown += dgvColumns_CellMouseDown;

            //cboStyle.SelectedIndexChanged += cboStyle_SelectedIndexChanged; // Subscribe to SelectedIndexChanged event
            //cboCustomer.SelectedIndexChanged += cboCustomer_SelectedIndexChanged; // Subscribe to SelectedIndexChanged even


        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //object[] rowData1 = new object[] { " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " " };

            dgvColumns.RowHeadersVisible = true;

            // Set RowHeadersWidthSizeMode to enable automatic resizing
            dgvColumns.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

            // Handle the RowPostPaint event to draw row numbers
            dgvColumns.RowPostPaint += dgvColumns_RowPostPaint;

        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            dgvColumns.Rows.Add(selectedStyle, selectedCustomer, "DATED", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
            dgvColumns.Rows.Add(selectedStyle, selectedCustomer, "STYLE", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
            dgvColumns.Rows.Add(selectedStyle, selectedCustomer, "CUSTOMER BRAND", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
            dgvColumns.Rows.Add(selectedStyle, selectedCustomer, "COUNTRY OF ORIGIN", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
            dgvColumns.Rows.Add(selectedStyle, selectedCustomer, "FACTORY", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
            dgvColumns.Rows.Add(selectedStyle, selectedCustomer, "DESCRIPTION", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
            dgvColumns.Rows.Add(selectedStyle, selectedCustomer, "SEASON", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");

            DataGridViewRow fabricRow1 = new DataGridViewRow();
            fabricRow1.CreateCells(dgvColumns, selectedStyle, selectedCustomer, "FABRIC", "CODE", "", "", "", "", "", "", "UNIT", "CIF BKK", "CONSUMPTION/PC", "AMOUNT/USD");


            for (int i = 0; i < fabricRow1.Cells.Count; i++)
            {
                fabricRow1.Cells[i].ReadOnly = true;
            }
            dgvColumns.Rows.Add(fabricRow1);

            dgvColumns.Rows.Add(selectedStyle, selectedCustomer, "A) BODY", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
            dgvColumns.Rows.Add(selectedStyle, selectedCustomer, "B) POCKET LINING", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
            dgvColumns.Rows.Add(selectedStyle, selectedCustomer, "C) NECK BINDING", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");



            DataGridViewRow fabricRow2 = new DataGridViewRow();
            fabricRow2.CreateCells(dgvColumns, selectedStyle, selectedCustomer, "TRIMS", "CODE", "PLACEMENT", "", "", "", "", "", "", "UNIT/PC", "QUANTITY", "AMOUNT/USD");


            for (int i = 0; i < fabricRow2.Cells.Count; i++)
            {
                fabricRow2.Cells[i].ReadOnly = true;
            }

            dgvColumns.Rows.Add(fabricRow2);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Create a SQL command to retrieve the maximum RowValue from the database
                string selectMaxRowQuery = "SELECT ISNULL(MAX([Row]), 0) FROM TestTableExcel;";

                SqlCommand selectMaxRowCommand = new SqlCommand(selectMaxRowQuery, connection);

                // Get the maximum RowValue from the database
                int maxRowValue = Convert.ToInt32(selectMaxRowCommand.ExecuteScalar());

                // Set the new RowValue to be one greater than the maximum RowValue in the database
                int newRowValue = maxRowValue + 1;

                // Create a SQL command to insert data
                string insertQuery = @"
            INSERT INTO TestTableExcel ([Row], Style, Customer, F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12)
            VALUES (@RowValue, @Style, @Customer, @F1, @F2, @F3, @F4, @F5, @F6, @F7, @F8, @F9, @F10, @F11, @F12);";

                SqlCommand insertCommand = new SqlCommand(insertQuery, connection);


                insertCommand.Parameters.Add("@RowValue", SqlDbType.Int).Value = newRowValue;
                insertCommand.Parameters.Add("@Style", SqlDbType.VarChar);
                insertCommand.Parameters.Add("@Customer", SqlDbType.VarChar);
                insertCommand.Parameters.Add("@F1", SqlDbType.VarChar);
                insertCommand.Parameters.Add("@F2", SqlDbType.VarChar);
                insertCommand.Parameters.Add("@F3", SqlDbType.VarChar);
                insertCommand.Parameters.Add("@F4", SqlDbType.VarChar);
                insertCommand.Parameters.Add("@F5", SqlDbType.VarChar);
                insertCommand.Parameters.Add("@F6", SqlDbType.VarChar);
                insertCommand.Parameters.Add("@F7", SqlDbType.VarChar);
                insertCommand.Parameters.Add("@F8", SqlDbType.VarChar);
                insertCommand.Parameters.Add("@F9", SqlDbType.VarChar);
                insertCommand.Parameters.Add("@F10", SqlDbType.VarChar);
                insertCommand.Parameters.Add("@F11", SqlDbType.VarChar);
                insertCommand.Parameters.Add("@F12", SqlDbType.VarChar);


                bool success = true;


                foreach (DataGridViewRow row in dgvColumns.Rows)
                {

                    string style = row.Cells[0].Value?.ToString() ?? "";
                    string customer = row.Cells[1].Value?.ToString() ?? "";
                    string f1 = row.Cells[2].Value?.ToString() ?? "";
                    string f2 = row.Cells[3].Value?.ToString() ?? "";
                    string f3 = row.Cells[4].Value?.ToString() ?? "";
                    string f4 = row.Cells[5].Value?.ToString() ?? "";
                    string f5 = row.Cells[6].Value?.ToString() ?? "";
                    string f6 = row.Cells[7].Value?.ToString() ?? "";
                    string f7 = row.Cells[8].Value?.ToString() ?? "";
                    string f8 = row.Cells[9].Value?.ToString() ?? "";
                    string f9 = row.Cells[10].Value?.ToString() ?? "";
                    string f10 = row.Cells[11].Value?.ToString() ?? "";
                    string f11 = row.Cells[12].Value?.ToString() ?? "";
                    string f12 = row.Cells[13].Value?.ToString() ?? "";


                    insertCommand.Parameters["@Style"].Value = style;
                    insertCommand.Parameters["@Customer"].Value = customer;
                    insertCommand.Parameters["@F1"].Value = f1;
                    insertCommand.Parameters["@F2"].Value = f2;
                    insertCommand.Parameters["@F3"].Value = f3;
                    insertCommand.Parameters["@F4"].Value = f4;
                    insertCommand.Parameters["@F5"].Value = f5;
                    insertCommand.Parameters["@F6"].Value = f6;
                    insertCommand.Parameters["@F7"].Value = f7;
                    insertCommand.Parameters["@F8"].Value = f8;
                    insertCommand.Parameters["@F9"].Value = f9;
                    insertCommand.Parameters["@F10"].Value = f10;
                    insertCommand.Parameters["@F11"].Value = f11;
                    insertCommand.Parameters["@F12"].Value = f12;


                    try
                    {
                        insertCommand.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        success = false;
                        MessageBox.Show("Error: " + ex.Message);
                        break;
                    }
                }

                // Check if all insertions were successful
                if (success)
                {
                    MessageBox.Show("All data inserted successfully.");
                }
                else
                {
                    MessageBox.Show("Some data could not be inserted. Please check and try again.");
                }
            }
        }

        private void dgvColumns_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && e.RowIndex != -1) // Right-click on a row
            {
                dgvColumns.ClearSelection();
                dgvColumns.Rows[e.RowIndex].Selected = true;

                ContextMenuStrip menu = new ContextMenuStrip();
                ToolStripMenuItem addItem = new ToolStripMenuItem("Insert Row Above");
                ToolStripMenuItem deleteItem = new ToolStripMenuItem("Delete Row");
                addItem.Click += InsertRow_Click;
                deleteItem.Click += DeleteRow_Click;
                menu.Items.Add(addItem);
                menu.Items.Add(deleteItem);

                menu.Show(Cursor.Position);
            }
        }

        private void InsertRow_Click(object sender, EventArgs e)
        {
            int rowIndex = dgvColumns.SelectedRows[0].Index;
            dgvColumns.Rows.Insert(rowIndex, new DataGridViewRow());
        }

        private void DeleteRow_Click(object sender, EventArgs e)
        {
            if (dgvColumns.SelectedRows.Count > 0)
            {
                if (dgvColumns.SelectedRows[0].IsNewRow)
                {

                    return;
                }

                dgvColumns.Rows.RemoveAt(dgvColumns.SelectedRows[0].Index);
            }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            dgvColumns.Rows.Clear();
        }

        private void cboStyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            //string selectedValue = cboStyle.SelectedItem.ToString();
            selectedStyle = cboStyle.SelectedItem.ToString();
            //MessageBox.Show("Selected Item: " + selectedStyle);
        }

        private void cboCustomer_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedCustomer = cboCustomer.SelectedItem.ToString();
            //MessageBox.Show("Selected Item: " + selectedCustomer);
        }


        private void PopulateComboBox()
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    string query = "SELECT * FROM CUSTOMER";
                    SqlCommand command = new SqlCommand(query, connection);
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    cboStyle.Items.Clear();
                    while (reader.Read())
                    {
                        cboStyle.Items.Add(reader.GetString(1));
                    }
                    reader.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }



            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    string query = "SELECT * FROM AGENTS";
                    SqlCommand command = new SqlCommand(query, connection);
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    cboCustomer.Items.Clear();
                    while (reader.Read())
                    {
                        cboCustomer.Items.Add(reader.GetString(0));
                    }
                    reader.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void dgvColumns_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.RowIndex != -1 && e.ColumnIndex != -1)
            //{
            //    DataGridViewCell clickedCell = dgvColumns.Rows[e.RowIndex].Cells[e.ColumnIndex];
            //    string cellValue = clickedCell.Value?.ToString(); // Get the value of the clicked cell
            //    if (cellValue != null)
            //    {
            //        textBox1.Text = cellValue; // Display the cell value in a TextBox
            //    }
            //}

            if (dgvColumns != null && textBox1 != null) // Check if dgvColumns and textBox1 are not null
            {
                if (e.RowIndex != -1 && e.ColumnIndex != -1)
                {
                    // Check if the clicked cell is within the bounds of the DataGridView
                    if (e.RowIndex < dgvColumns.Rows.Count && e.ColumnIndex < dgvColumns.Columns.Count)
                    {
                        // หาชื่อของคอลัมน์ที่ถูกคลิก
                        string columnName = dgvColumns.Columns[e.ColumnIndex].HeaderText;

                        // หาค่าที่ถูกคลิก
                        string cellValue = dgvColumns.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString();

                        // Get the row number
                        int rowNumber = e.RowIndex + 1;

                        // แสดงชื่อและค่าที่ถูกคลิกใน TextBox
                        textBox1.Text = "Row Number: " + rowNumber + ", Column Name: " + columnName + ", Cell Value: " + cellValue;
                    }
                }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //string formula = textBox1.Text.Trim(); // รับสูตรจาก TextBox

            //if (!string.IsNullOrEmpty(formula))
            //{
            //    try
            //    {
            //        // ดึงเซลล์ที่เลือกใน DataGridView
            //        DataGridViewCell selectedCell = dgvColumns.CurrentCell;

            //        // ตรวจสอบว่ามีเซลล์ที่เลือกหรือไม่
            //        if (selectedCell != null)
            //        {
            //            DataTable table = new DataTable();
            //            // คำนวณผลลัพธ์ของสูตร
            //            object result = table.Compute(formula, "");

            //            // แสดงผลลัพธ์ในเซลล์ที่เลือก
            //            selectedCell.Value = result.ToString();
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show("Error: " + ex.Message);
            //    }
            //}


        }

        private void dgvColumns_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            string rowNumber = (e.RowIndex + 1).ToString();

            // Create a rectangle to position the row number
            Rectangle rowBounds = new Rectangle(e.RowBounds.Location.X,
                                                e.RowBounds.Location.Y,
                                                dgvColumns.RowHeadersWidth - 4,
                                                e.RowBounds.Height);

            // Center the row number vertically
            TextRenderer.DrawText(e.Graphics,
                                rowNumber,
                                dgvColumns.RowHeadersDefaultCellStyle.Font,
                                rowBounds,
                                dgvColumns.RowHeadersDefaultCellStyle.ForeColor,
                                TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }
    }
}
