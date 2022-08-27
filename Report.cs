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
            if (Form1.kieu_bao_cao == 1)
            {
                SqlCommand cmd = new SqlCommand("report_by_huyen", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@ngay", Form1.day));
                da = new SqlDataAdapter(cmd);
                System.Data.DataTable dt1 = new System.Data.DataTable();
                da.Fill(dt1);
                DataSet ds = new DataSet();
                da.Fill(ds);
                reportViewer1.ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Local;
                reportViewer1.LocalReport.ReportPath = "Report1.rdlc";
                if (ds.Tables[0].Rows.Count > 0)
                {
                    ReportDataSource rds = new ReportDataSource();
                    rds.Name = "DataSet1";
                    rds.Value = ds.Tables[0];
                    reportViewer1.LocalReport.DataSources.Clear();
                    reportViewer1.LocalReport.DataSources.Add(rds);
                    reportViewer1.RefreshReport();
                }
            }
            if (Form1.kieu_bao_cao == 2)
            {
                SqlCommand cmd = new SqlCommand("report_by_huyen_all", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da2 = new SqlDataAdapter(cmd);
                System.Data.DataTable dt2 = new System.Data.DataTable();
                da2.Fill(dt2);
                DataSet ds = new DataSet();
                da.Fill(ds);
                reportViewer1.ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Local;
                reportViewer1.LocalReport.ReportPath = "Report_Duyet.rdlc";
                if (ds.Tables[0].Rows.Count > 0)
                {
                    ReportDataSource rds = new ReportDataSource();
                    rds.Name = "DataSet2";
                    rds.Value = ds.Tables[0];
                    reportViewer1.LocalReport.DataSources.Clear();
                    reportViewer1.LocalReport.DataSources.Add(rds);
                    reportViewer1.RefreshReport();
                }
            }
        }
    }
}
