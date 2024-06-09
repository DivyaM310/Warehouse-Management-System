using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
namespace WarehouseManagementSystem
{
    public partial class frmMainMenu : Form
    {
        SqlDataReader rdr = null;
        SqlConnection con = null;
        SqlCommand cmd = null;
        ConnectionString cs = new ConnectionString();
        public frmMainMenu()
        {
            InitializeComponent();
        }

        private void customerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmCustomers frm = new frmCustomers();
            frm.Show();
        }

        private void registrationToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            frmRegistration frm = new frmRegistration();
            frm.Show();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmAbout frm = new frmAbout();
            frm.Show();
        }

        private void registrationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmRegistration frm = new frmRegistration();
            frm.Show();
        }

        private void profileEntryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmCustomers frm = new frmCustomers();
            frm.Show();
        }

        private void productToolStripMenuItem_Click(object sender, EventArgs e)
        {

            frmProduct frm = new frmProduct();
            frm.Show();
        }

        private void notepadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("Notepad.exe");
        }

        private void calculatorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("Calc.exe");
        }

        private void wordpadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("Wordpad.exe");
        }

        private void taskManagerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("TaskMgr.exe");
        }

        private void mSWordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("Winword.exe");
        }

        private void categoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmCategory frm = new frmCategory();
            frm.Show();
        }

        private void companyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmSubCategory frm = new frmSubCategory();
            frm.Show();
        }

        private void customersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmCustomersRecord frm = new frmCustomersRecord();
            frm.Show();
        }

        private void logOutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmCategory o1 = new frmCategory();
            o1.Hide();
            frmSubCategory o2 = new frmSubCategory();
            o2.Hide();
            frmProduct o3 = new frmProduct();
            o3.Hide();
            frmRegisteredUsersDetails o4 = new frmRegisteredUsersDetails();
            o4.Hide();
            frmRegistration o5 = new frmRegistration();
            o5.Hide();
            frmStockRecord o6 = new frmStockRecord();
            o6.Hide();
            frmCustomersRecord o7 = new frmCustomersRecord();
            o7.Hide();
            frmSuppliersRecord o8 = new frmSuppliersRecord();
            o8.Hide();
            frmProductsRecord2 o9 = new frmProductsRecord2();
            o9.Hide();
            frmSalesRecord2 o10 = new frmSalesRecord2();
            o10.Hide();
            frmLogin frm = new frmLogin();
            frm.Show();
            frm.txtUserName.Text = "";
            frm.txtPassword.Text = "";
            frm.ProgressBar1.Visible = false;
            frm.txtUserName.Focus();
        }

        private void frmMainMenu_Load(object sender, EventArgs e)
        {
            if (lblUserType.Text.Trim()== "Admin")
                    {
                       
                        this.masterEntryToolStripMenuItem.Enabled=true;
                        this.usersToolStripMenuItem.Enabled=true;
                        this.customerToolStripMenuItem1.Enabled=true;
                        this.suppliersToolStripMenuItem.Enabled=true;
                        this.productsToolStripMenuItem.Enabled=true;
                        this.recordsToolStripMenuItem.Enabled=true;
                        this.registrationToolStripMenuItem.Enabled=true;
                        //this.databaseToolStripMenuItem.Enabled=true;
                        this.customerToolStripMenuItem.Enabled=true;
                        this.supplierToolStripMenuItem.Enabled=true;
                        this.productToolStripMenuItem.Enabled=true;
                        this.stockToolStripMenuItem.Enabled=true;
                        this.invoiceToolStripMenuItem.Enabled = true;
   
         
                    }
                    if (lblUserType.Text.Trim() == "Sales Person")
                    {
                        this.masterEntryToolStripMenuItem.Enabled = false;
                        this.usersToolStripMenuItem.Enabled = false;
                        this.customerToolStripMenuItem1.Enabled = true;
                        this.suppliersToolStripMenuItem.Enabled = false;
                        this.productsToolStripMenuItem.Enabled = false;
                        this.recordsToolStripMenuItem.Enabled = false;
                        this.registrationToolStripMenuItem.Enabled = false;
                        //this.databaseToolStripMenuItem.Enabled = false;
                        this.customerToolStripMenuItem.Enabled = true;
                        this.supplierToolStripMenuItem.Enabled = false;
                        this.productToolStripMenuItem.Enabled = false;
                        this.stockToolStripMenuItem.Enabled = false;
                        this.invoiceToolStripMenuItem.Enabled = true;
               
                    }
                    if (lblUserType.Text.Trim() == "Warehouse Worker")
                    {
                        this.masterEntryToolStripMenuItem.Enabled = false;
                        this.usersToolStripMenuItem.Enabled = false;
                        this.customerToolStripMenuItem1.Enabled = false;
                        this.suppliersToolStripMenuItem.Enabled = false;
                        this.productsToolStripMenuItem.Enabled = false;
                        this.recordsToolStripMenuItem.Enabled = false;
                        this.registrationToolStripMenuItem.Enabled = false;
                        //this.databaseToolStripMenuItem.Enabled = false;
                        this.customerToolStripMenuItem.Enabled = false;
                        this.supplierToolStripMenuItem.Enabled = false;
                        this.productToolStripMenuItem.Enabled = false;
                        this.stockToolStripMenuItem.Enabled = false;
                        this.invoiceToolStripMenuItem.Enabled = false;
                   
                    }
                
                       if (lblUserType.Text.Trim() == "Warehouse Manager")
                    {
                        this.masterEntryToolStripMenuItem.Enabled = false;
                        this.usersToolStripMenuItem.Enabled = false;
                        this.customerToolStripMenuItem1.Enabled = false;
                        this.suppliersToolStripMenuItem.Enabled = false;
                        this.productsToolStripMenuItem.Enabled = false;
                        this.recordsToolStripMenuItem.Enabled = false;
                        this.registrationToolStripMenuItem.Enabled = false;
                       // this.databaseToolStripMenuItem.Enabled = false;
                        this.customerToolStripMenuItem.Enabled = false;
                        this.supplierToolStripMenuItem.Enabled = true;
                        this.productToolStripMenuItem.Enabled = true;
                        this.stockToolStripMenuItem.Enabled = true;
                        this.invoiceToolStripMenuItem.Enabled = false;
                  
                    }
                        ToolStripStatusLabel4.Text = System.DateTime.Now.ToString();
                        GetData();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            ToolStripStatusLabel4.Text = System.DateTime.Now.ToString();
        }

        private void productsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            frmProduct frm = new frmProduct();
            frm.Show();
        }

        private void productsToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            frmProductsRecord2 frm = new frmProductsRecord2();
            frm.Show();
        }

     
        private void stockToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmStock frm = new frmStock();
            frm.lblUser.Text = lblUser.Text;
            frm.lblUserType.Text = lblUserType.Text;
            frm.Show();
        }

        private void stockToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmStock frm = new frmStock();
            frm.lblUser.Text = lblUser.Text;
            frm.lblUserType.Text = lblUserType.Text;
            frm.Show();
        }

        public void GetData()
        {
            try
            {
                con = new SqlConnection(cs.DBConn);
                con.Open();
                String sql = "SELECT Product.ProductID,ProductName,Features,Price,sum(Quantity),sum(Price*Quantity) from Temp_Stock,Product where Temp_Stock.ProductID=Product.ProductID group by Product.productID,productname,Price,Features,Quantity having(Quantity>0)  order by ProductName";
                cmd = new SqlCommand(sql, con);
                rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                dataGridView1.Rows.Clear();
                while (rdr.Read() == true)
                {
                    dataGridView1.Rows.Add(rdr[0], rdr[1], rdr[2], rdr[3], rdr[4], rdr[5]);
                }
                foreach (DataGridViewRow r in this.dataGridView1.Rows)
                {
                    if (Convert.ToInt32(r.Cells[4].Value) < 10)
                    {
                        r.DefaultCellStyle.BackColor = Color.Red;
                    }
                }
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

   
        private void invoiceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmSales frm = new frmSales();
            frm.lblUser.Text = lblUser.Text;
            frm.lblUserType.Text = lblUserType.Text;
            frm.Show();
        }

        private void salesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmSales frm = new frmSales();
            frm.lblUser.Text = lblUser.Text;
            frm.lblUserType.Text = lblUserType.Text;
            frm.Show();
        }

        private void salesToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            frmSalesRecord2 frm = new frmSalesRecord2();
            frm.Show();
        }

        private void loginDetailsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmLoginDetails frm = new frmLoginDetails();
            frm.Show();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try{
            con = new SqlConnection(cs.DBConn);
                con.Open();
                String sql = "SELECT Product.ProductID,ProductName,Features,Price,sum(Quantity),sum(Price*Quantity) from Temp_Stock,Product where Temp_Stock.ProductID=Product.ProductID and ProductName like '" + txtProductName.Text + "%' group by product.ProductID,productname,Price,Features,Quantity having(quantity>0) order by ProductName";
                cmd = new SqlCommand(sql, con);
                rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                dataGridView1.Rows.Clear();
                while (rdr.Read() == true)
                {
                    dataGridView1.Rows.Add(rdr[0], rdr[1], rdr[2], rdr[3], rdr[4], rdr[5]);
                }
                foreach (DataGridViewRow r in this.dataGridView1.Rows)
                {
                if (Convert.ToInt32(r.Cells[4].Value) < 10)
                {
                    r.DefaultCellStyle.BackColor = Color.Red;
                }
            }
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void frmMainMenu_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
            
        }

        private void profileEntryToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            frmSuppliers frm = new frmSuppliers();
            frm.Show();
        }

        private void supplierToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmSuppliers frm = new frmSuppliers();
            frm.Show();
        }

        private void suppliersToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            frmSuppliersRecord frm = new frmSuppliersRecord();
            frm.Show();
        }

        private void stockToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            frmStockRecord frm = new frmStockRecord();
            frm.Show();
        }

        private void backupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                timer2.Enabled = true;
                if ((!System.IO.Directory.Exists("C:\\DBBackup")))
                {
                    System.IO.Directory.CreateDirectory("C:\\DBBackup");
                }
                string destdir = "C:\\DBBackup\\pos_db " + DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss") + ".bak";
                con = new SqlConnection(cs.DBConn);
                con.Open();
                string cb = "backup database [" + System.Windows.Forms.Application.StartupPath + "\\pos_db.mdf] to disk='" + destdir + "'with init,stats=10";
                cmd = new SqlCommand(cb);
                cmd.Connection = con;
                cmd.ExecuteReader();
                con.Close();
                MessageBox.Show("Successfully performed", "Database Backup", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void restoreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                var _with1 = openFileDialog1;
                _with1.Filter = ("DB Backup File|*.bak;");
                _with1.FilterIndex = 4;
                //Clear the file name
                openFileDialog1.FileName = "";

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Cursor = Cursors.WaitCursor;
                    timer2.Enabled = true;
                    SqlConnection.ClearAllPools();
                    con = new SqlConnection(cs.DBConn);
                    con.Open();
                    string cb = "USE Master ALTER DATABASE [" + System.Windows.Forms.Application.StartupPath + "\\pos_db.mdf] SET Single_User WITH Rollback Immediate Restore database [" + System.Windows.Forms.Application.StartupPath + "\\pos_db.mdf] FROM disk='" + openFileDialog1.FileName + "' WITH REPLACE ALTER DATABASE [" + System.Windows.Forms.Application.StartupPath + "\\pos_db.mdf] SET Multi_User ";
                    cmd = new SqlCommand(cb);
                    cmd.Connection = con;
                    cmd.ExecuteReader();
                    MessageBox.Show("Successfully performed", "Database Restore", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GetData();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            Cursor = Cursors.Default;
            timer2.Enabled = false;
        }

        private void btnGetData_Click(object sender, EventArgs e)
        {
 
            int rowsTotal = 0;
            int colsTotal = 0;
            int I = 0;
            int j = 0;
            int iC = 0;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Excel.Application xlApp = new Excel.Application();

            try
            {
                Excel.Workbook excelBook = xlApp.Workbooks.Add();
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelBook.Worksheets[1];
                xlApp.Visible = true;

                rowsTotal = dataGridView1.RowCount;
                colsTotal = dataGridView1.Columns.Count - 1;
                var _with1 = excelWorksheet;
                _with1.Cells.Select();
                _with1.Cells.Delete();
                for (iC = 0; iC <= colsTotal; iC++)
                {
                    _with1.Cells[1, iC + 1].Value = dataGridView1.Columns[iC].HeaderText;
                }
                for (I = 0; I <= rowsTotal - 1; I++)
                {
                    for (j = 0; j <= colsTotal; j++)
                    {
                        _with1.Cells[I + 2, j + 1].value = dataGridView1.Rows[I].Cells[j].Value;
                    }
                }
                _with1.Rows["1:1"].Font.FontStyle = "Bold";
                _with1.Rows["1:1"].Font.Size = 12;

                _with1.Cells.Columns.AutoFit();
                _with1.Cells.Select();
                _with1.Cells.EntireColumn.AutoFit();
                _with1.Cells[1, 1].Select();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                //RELEASE ALLOACTED RESOURCES
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                xlApp = null;
            }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            txtProductName.Text = "";
            GetData();
        }

        private void lowStockToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmLowStockReport frm = new frmLowStockReport();
            frm.ShowDialog();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
