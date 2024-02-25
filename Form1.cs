using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

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


            //cboStyle.SelectedIndexChanged += cboStyle_SelectedIndexChanged; // Subscribe to SelectedIndexChanged event
            //cboCustomer.SelectedIndexChanged += cboCustomer_SelectedIndexChanged; // Subscribe to SelectedIndexChanged even


        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //object[] rowData1 = new object[] { " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " " };

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
            fabricRow1.CreateCells(dgvColumns, selectedStyle, selectedCustomer, "FABRIC", "CODE", "", "", "", "", "", "", "", "", "UNIT", "CIF BKK", "CONSUMPTION/PC", "AMOUNT/USD");

            // Set all cells in fabricRow1 as read-only
            for (int i = 0; i < fabricRow1.Cells.Count; i++)
            {
                fabricRow1.Cells[i].ReadOnly = true;
            }

            // Add fabricRow1 to the DataGridView
            dgvColumns.Rows.Add(fabricRow1);

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

                // Add parameters to the insert command
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

                // Track if any insert operation failed
                bool success = true;

                // Iterate through DataGridView rows and execute the insert command for each row
                foreach (DataGridViewRow row in dgvColumns.Rows)
                {
                    // Check if the cell's value is not null before accessing it
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

                    // Set parameter values
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

                    // Execute the command
                    try
                    {
                        insertCommand.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        // If an error occurs during insertion, set success to false and break the loop
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
                        cboStyle.Items.Add(reader.GetString(2));
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
                        cboCustomer.Items.Add(reader.GetString(1));
                    }
                    reader.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }


    }
}
