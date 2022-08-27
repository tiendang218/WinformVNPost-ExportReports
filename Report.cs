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
        {   switch(Form1.kieu_bao_cao)
            {
                case 1:
                    SqlCommand cmd1 = new SqlCommand("report_by_huyen", cn);
                    cmd1.CommandType = CommandType.StoredProcedure;
                    cmd1.Parameters.Add(new SqlParameter("@ngay", Form1.day));
                    da = new SqlDataAdapter(cmd1);
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    reportViewer1.Reset();
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
                    break;
                case 2:
                    SqlCommand cmd2 = new SqlCommand("report_by_huyen_all", cn);
                    cmd2.CommandType = CommandType.StoredProcedure;
                    SqlDataAdapter da2 = new SqlDataAdapter(cmd2);
                    DataSet ds1 = new DataSet();
                    da2.Fill(ds1);
                    reportViewer1.Reset();
                    reportViewer1.ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Local;
                    reportViewer1.LocalReport.ReportPath = "ReportALL.rdlc";
                    if (ds1.Tables[0].Rows.Count > 0)
                    {
                        ReportDataSource rds1 = new ReportDataSource();
                        rds1.Name = "DataSet2";
                        rds1.Value = ds1.Tables[0];
                        reportViewer1.LocalReport.DataSources.Clear();
                        reportViewer1.LocalReport.DataSources.Add(rds1);
                        reportViewer1.RefreshReport();
                    }
                    break;
                case 3:
                    SqlCommand cmd3 = new SqlCommand("report_by_huyen_duyet", cn);
                    cmd3.CommandType = CommandType.StoredProcedure;
                    cmd3.Parameters.Add(new SqlParameter("@ngay", Form1.day));
                    SqlDataAdapter da3 = new SqlDataAdapter(cmd3);
                    DataSet ds2 = new DataSet();
                    da3.Fill(ds2);
                    reportViewer1.Reset();
                    reportViewer1.ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Local;
                    reportViewer1.LocalReport.ReportPath = "Report_ngay_duyet.rdlc";
                    if (ds2.Tables[0].Rows.Count > 0)
                    {
                        ReportDataSource rds2 = new ReportDataSource();
                        rds2.Name = "DataSet3";
                        rds2.Value = ds2.Tables[0];
                        reportViewer1.LocalReport.DataSources.Clear();
                        reportViewer1.LocalReport.DataSources.Add(rds2);
                        reportViewer1.RefreshReport();
                    }
                    break;
                case 4:
                    SqlCommand cmd4 = new SqlCommand("report_by_huyen_hen", cn);
                    cmd4.CommandType = CommandType.StoredProcedure;
                    cmd4.Parameters.Add(new SqlParameter("@ngay", Form1.day));
                    SqlDataAdapter da4 = new SqlDataAdapter(cmd4);
                    DataSet ds3 = new DataSet();
                    da4.Fill(ds3);
                    reportViewer1.Reset();
                    reportViewer1.ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Local;
                    reportViewer1.LocalReport.ReportPath = "Report_ngay_hen.rdlc";
                    if (ds3.Tables[0].Rows.Count > 0)
                    {
                        ReportDataSource rds3 = new ReportDataSource();
                        rds3.Name = "DataSet4";
                        rds3.Value = ds3.Tables[0];
                        reportViewer1.LocalReport.DataSources.Clear();
                        reportViewer1.LocalReport.DataSources.Add(rds3);
                        reportViewer1.RefreshReport();
                    }
                    break;
                default:
                    break;
            }
        }
    }
}
