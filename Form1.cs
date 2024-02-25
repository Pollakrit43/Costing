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
        private string selectedType;

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

            dgvColumns.CellValueChanged += dgvColumns_CellValueChanged;

        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            // Check if the selected style and customer are not empty
            if (!string.IsNullOrEmpty(selectedStyle) && !string.IsNullOrEmpty(selectedCustomer) && !string.IsNullOrEmpty(tb2.Text))
            {
                dgvColumns.Rows.Add(selectedType, selectedStyle, selectedCustomer, "DATED", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, "STYLE", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, "CUSTOMER BRAND", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, "COUNTRY OF ORIGIN", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, "FACTORY", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, "DESCRIPTION", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, "SEASON", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");

                DataGridViewRow fabricRow1 = new DataGridViewRow();
                fabricRow1.CreateCells(dgvColumns, selectedType,selectedStyle, selectedCustomer, "FABRIC", "CODE", "", "", "", "", "", "", "UNIT", "CIF BKK", "CONSUMPTION/PC", "AMOUNT/USD");

                for (int i = 0; i < fabricRow1.Cells.Count; i++)
                {
                    fabricRow1.Cells[i].ReadOnly = true;
                }
                dgvColumns.Rows.Add(fabricRow1);

                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, "A) BODY", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, "B) POCKET LINING", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, "C) NECK BINDING", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");

                DataGridViewRow fabricRow2 = new DataGridViewRow();
                fabricRow2.CreateCells(dgvColumns, selectedType, selectedStyle, selectedCustomer, "TRIMS", "CODE", "PLACEMENT", "", "", "", "", "", "", "UNIT/PC", "QUANTITY", "AMOUNT/USD");

                for (int i = 0; i < fabricRow2.Cells.Count; i++)
                {
                    fabricRow2.Cells[i].ReadOnly = true;
                }

                dgvColumns.Rows.Add(fabricRow2);


                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, "SCREEN PRINT", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, "LOGO HEAT TRANSFER", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, "THREAD", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");


                DataGridViewRow fabricRow3 = new DataGridViewRow();
                fabricRow3.CreateCells(dgvColumns, selectedType,selectedStyle, selectedCustomer, "LABELLING", "", "", "", "", "", "", "", "", "UNIT/PC", "QUANTITY", "AMOUNT/USD");

                for (int i = 0; i < fabricRow3.Cells.Count; i++)
                {
                    fabricRow3.Cells[i].ReadOnly = true;
                }

                dgvColumns.Rows.Add(fabricRow3);

                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");


                DataGridViewRow fabricRow4 = new DataGridViewRow();
                fabricRow4.CreateCells(dgvColumns, selectedType, selectedStyle, selectedCustomer, "PACKING & FINISHING", "", "", "", "", "", "", "", "", "UNIT/PC", "QUANTITY", "AMOUNT/USD");

                for (int i = 0; i < fabricRow4.Cells.Count; i++)
                {
                    fabricRow4.Cells[i].ReadOnly = true;
                }

                dgvColumns.Rows.Add(fabricRow4);


                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ");
                dgvColumns.Rows.Add(selectedType,selectedStyle, selectedCustomer, " ", " ", " ", " ", " ", " ", "Handling Charge", " ", " ", " ", " ", " ", " ", " ");

                DataGridViewRow fabricRow5 = new DataGridViewRow();
                fabricRow5.CreateCells(dgvColumns, selectedType, selectedStyle, selectedCustomer, "TOTAL MATERIAL COST", "", "", "", "", "", "", "", "", "", "", "");

                for (int i = 0; i < fabricRow5.Cells.Count; i++)
                {
                    fabricRow5.Cells[i].ReadOnly = true;
                }

                dgvColumns.Rows.Add(fabricRow5);

            }
            else
            {
                MessageBox.Show("Please select a style, customer, and enter a value in the textbox.");
            }
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
            INSERT INTO TestTableExcel ([Row],Type, Style, Customer, F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12)
            VALUES (@RowValue,@Type, @Style, @Customer, @F1, @F2, @F3, @F4, @F5, @F6, @F7, @F8, @F9, @F10, @F11, @F12);";

                SqlCommand insertCommand = new SqlCommand(insertQuery, connection);


                insertCommand.Parameters.Add("@RowValue", SqlDbType.Int).Value = newRowValue;
                insertCommand.Parameters.Add("@Type", SqlDbType.VarChar);
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
                    string type = row.Cells[0].Value?.ToString() ?? "";
                    string style = row.Cells[1].Value?.ToString() ?? "";
                    string customer = row.Cells[2].Value?.ToString() ?? "";
                    string f1 = row.Cells[3].Value?.ToString() ?? "";
                    string f2 = row.Cells[4].Value?.ToString() ?? "";
                    string f3 = row.Cells[5].Value?.ToString() ?? "";
                    string f4 = row.Cells[6].Value?.ToString() ?? "";
                    string f5 = row.Cells[7].Value?.ToString() ?? "";
                    string f6 = row.Cells[8].Value?.ToString() ?? "";
                    string f7 = row.Cells[9].Value?.ToString() ?? "";
                    string f8 = row.Cells[10].Value?.ToString() ?? "";
                    string f9 = row.Cells[11].Value?.ToString() ?? "";
                    string f10 = row.Cells[12].Value?.ToString() ?? "";
                    string f11 = row.Cells[13].Value?.ToString() ?? "";
                    string f12 = row.Cells[14].Value?.ToString() ?? "";


                    insertCommand.Parameters["@Type"].Value = type;
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
                ToolStripMenuItem addItemAbove = new ToolStripMenuItem("Insert Row Above");
                ToolStripMenuItem addItemBelow = new ToolStripMenuItem("Insert Row Below");
                ToolStripMenuItem deleteItem = new ToolStripMenuItem("Delete Row");
                addItemAbove.Click += InsertRowAbove_Click;
                addItemBelow.Click += InsertRowBelow_Click;
                deleteItem.Click += DeleteRow_Click;
                menu.Items.Add(addItemAbove);
                menu.Items.Add(addItemBelow);
                menu.Items.Add(deleteItem);

                menu.Show(Cursor.Position);
            }
        }

        private void InsertRowAbove_Click(object sender, EventArgs e)
        {
            int rowIndex = dgvColumns.SelectedRows[0].Index;
            dgvColumns.Rows.Insert(rowIndex, new DataGridViewRow());
        }

        private void InsertRowBelow_Click(object sender, EventArgs e)
        {
            int rowIndex = dgvColumns.SelectedRows[0].Index;
            dgvColumns.Rows.Insert(rowIndex + 1, new DataGridViewRow());
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
            textBox1.Clear();
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

        private void cboType_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedType = cboType.SelectedItem.ToString();
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


            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    string query = "SELECT * FROM ORDERS";
                    SqlCommand command = new SqlCommand(query, connection);
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    cboType.Items.Clear();
                    while (reader.Read())
                    {
                        cboType.Items.Add(reader.GetString(5));
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
                        textBox1.Text = rowNumber + "-" + columnName + " Cell Value: " + cellValue;
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

        private void dgvColumns_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            // Check if the changed cell is in Row 3
            //if (e.RowIndex == 2)
            //{
            //    decimal sum = 1; // Initialize sum to 1 instead of 0

            //    // Loop through each cell in the row
            //    for (int columnIndex = 0; columnIndex < dgvColumns.Columns.Count; columnIndex++)
            //    {
            //        // Check if the current cell is not in F12 column (index 11)
            //        if (columnIndex != 13)
            //        {
            //            // Get the value from the cell
            //            string cellValue = dgvColumns.Rows[e.RowIndex].Cells[columnIndex].Value?.ToString();

            //            if (!string.IsNullOrEmpty(cellValue) && decimal.TryParse(cellValue, out decimal cellDecimalValue))
            //            {
            //                // Multiply the sum by the cell value
            //                sum *= cellDecimalValue;
            //            }
            //        }
            //    }

            //    // Multiply the accumulated sum by 1.05
            //    sum *= 1.05m;

            //    // Update the value in F12 cell (index 11) with the sum
            //    dgvColumns.Rows[e.RowIndex].Cells[13].Value = Math.Round(sum, 4).ToString("0.0000");
            //}
        }

        private void dgvColumns_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            // Check if the edited cell is in columns "F10" or "F11"
            if (e.ColumnIndex == 11 || e.ColumnIndex == 12)
            {
                // Get the value from tb2 textbox
                if (double.TryParse(tb2.Text, out double tb2Value))
                {
                    foreach (DataGridViewRow row in dgvColumns.Rows)
                    {
                        // Get the values from cells "F10" and "F11" of the current row
                        string f10Value = row.Cells["F10"].Value?.ToString();
                        string f11Value = row.Cells["F11"].Value?.ToString();

                        // Check if both values are not null or empty
                        if (!string.IsNullOrEmpty(f10Value) && !string.IsNullOrEmpty(f11Value))
                        {
                            // Parse the values to doubles
                            if (double.TryParse(f10Value, out double f10) && double.TryParse(f11Value, out double f11))
                            {
                                // Calculate the value for "F12" cell with a 5% increment
                                double result = f10 * f11 * tb2Value;

                                // Update the value in the "F12" cell
                                row.Cells["F12"].Value = result.ToString("0.0000");
                            }
                            else
                            {
                                // Handle invalid input in F10 or F11.
                            }
                        }
                    }
                }
                else
                {
                    // Handle invalid input in tb2.
                }
            }
        }

        private void tb2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
       (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

       
    }
}
