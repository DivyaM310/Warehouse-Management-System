using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
namespace WarehouseManagementSystem
{
    public partial class frmLowStockReport : Form
    {
   
        SqlConnection con = null;
        SqlCommand cmd = null;
        ConnectionString cs = new ConnectionString();
        public frmLowStockReport()
        {
            InitializeComponent();
        }

        private void frmLowStockReport_Load(object sender, EventArgs e)
        {
            try{
                
                rptLowStock rpt = new rptLowStock();
                //The report you created.
                cmd = new SqlCommand();
                SqlDataAdapter myDA = new SqlDataAdapter();
                POS_DBDataSet myDS = new POS_DBDataSet();
                //The DataSet you created.
                con = new SqlConnection(cs.DBConn);
                cmd.Connection = con;
                cmd.CommandText = "SELECT * from Temp_Stock,Product where Temp_Stock.ProductID=Product.ProductID and quantity < 10 order by ProductName";
                cmd.CommandType = CommandType.Text;
                myDA.SelectCommand = cmd;
                myDA.Fill(myDS, "Temp_Stock");
                myDA.Fill(myDS, "Product");
                rpt.SetDataSource(myDS);
                this.crystalReportViewer1.ReportSource = rpt;
                this.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

     
    }
}
