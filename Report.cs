using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XuatExcelApp
{
    public partial class Report : Form
    {
        SqlConnection cn = new SqlConnection(@"Data Source=KTNV-TIEN\SQLEXPRESS;Initial Catalog=ChungTuHCC;Integrated Security=True");
        SqlCommand cmd;
        SqlDataAdapter da;
        
        public Report()
        {
            InitializeComponent();
        }

        private void Report_Load(object sender, EventArgs e)
        {

            this.reportViewer1.RefreshReport();
            
        }

        private void reportViewer1_Load(object sender, EventArgs e)
        {
            cmd = new SqlCommand("report_by_huyen", cn);
            cmd.CommandType = CommandType.StoredProcedure; 
            cmd.Parameters.Add(new SqlParameter("@ngay", Form1.day));
            da = new SqlDataAdapter(cmd);
            System.Data.DataTable dt3 = new System.Data.DataTable();
            da.Fill(dt3);
            DataSet ds = new DataSet();
            da.Fill(ds);
            reportViewer1.ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Local;
            reportViewer1.LocalReport.ReportPath = "Report1.rdlc";
            if (ds.Tables[0].Rows.Count > 0)
            {
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSet3";
                rds.Value = ds.Tables[0];
                reportViewer1.LocalReport.DataSources.Clear();
                reportViewer1.LocalReport.DataSources.Add(rds);
                reportViewer1.RefreshReport();
            }
        }
    }
}
