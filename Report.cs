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
            //Khai báo câu lệnh SQL
            //String sql = "Select * from tblMatHang Where NgaySX >='" + dtpNgaySX.Value.ToString() + "'";
            //SqlConnection con = new SqlConnection();
            ////Truyền vào chuỗi kết nối tới cơ sở dữ liệu
            ////Gọi Application.StartupPath để lấy đường dẫn tới thư mục chứa file chạy chương trình 
            //con.ConnectionString = @"Data Source=.\SQLEXPRESS;AttachDbFilename=" + Application.StartupPath + @"\QLBanHang.mdf;Integrated Security=True;User Instance=True";
            //SqlDataAdapter adp = new SqlDataAdapter(sql, con);
            //DataSet ds = new DataSet();
            //adp.Fill(ds);


            //SqlCommand cmd = new SqlCommand("load_ma_huyen", cn);
            //cmd.CommandType = CommandType.StoredProcedure;
            //SqlDataAdapter da1 = new SqlDataAdapter(cmd);
            //System.Data.DataTable dt = new System.Data.DataTable();
            //da1.Fill(dt);
            

            //SqlCommand cmd2 = new SqlCommand("load_ct_huyen", cn);
            //cmd2.CommandType = CommandType.StoredProcedure;
            //foreach (DataRow dr in dt.Rows)
            //{
            //    foreach (DataColumn dc in dt.Columns)
            //    {
            //        var i = dr[dc].ToString();
            //        cmd2.Parameters.Add(new SqlParameter("@huyen", Int32.Parse(i)));
            //    }    
            //}
            //SqlDataAdapter da2 = new SqlDataAdapter(cmd2);
            //System.Data.DataTable dt2 = new System.Data.DataTable();
            //da2.Fill(dt2);


            SqlCommand cmd = new SqlCommand("report_by_huyen", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            da = new SqlDataAdapter(cmd);
            System.Data.DataTable dt3 = new System.Data.DataTable();
            da.Fill(dt3);
            DataSet ds = new DataSet();
            da.Fill(ds);
            //Khai báo chế độ xử lý báo cáo, trong trường hợp này lấy báo cáo ở local
            reportViewer1.ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Local;
            //Đường dẫn báo cáo
            reportViewer1.LocalReport.ReportPath = "Report1.rdlc";
            //Nếu có dữ liệu
            if (ds.Tables[0].Rows.Count > 0)
            {
                //Tạo nguồn dữ liệu cho báo cáo
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSet1";
                rds.Value = ds.Tables[0];
                //Xóa dữ liệu của báo cáo cũ trong trường hợp người dùng thực hiện câu truy vấn khác
                reportViewer1.LocalReport.DataSources.Clear();
                //Add dữ liệu vào báo cáo
                reportViewer1.LocalReport.DataSources.Add(rds);
                //Refresh lại báo cáo
                reportViewer1.RefreshReport();
            }
        }
    }
}
