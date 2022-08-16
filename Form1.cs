using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace XuatExcelApp
{
    public partial class Form1 : Form
    {
        SqlConnection cn;
        SqlCommand cmd;
        SqlDataAdapter da;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cn = new SqlConnection(@"Data Source=KTNV-TIEN\SQLEXPRESS;Initial Catalog=ChungTuHCC;Integrated Security=True");
            cn.Open();
            //bind data in data grid view  
            Get_All_ChungTu();
            load_tinh();
            load_loai_hoso();
            LoadHuyen(tinh_comboBox2.SelectedValue.ToString());
            LoadXa(huyen_comboBox3.SelectedValue.ToString());
            ////disable delete and update button on load  
            Sửa.Enabled = true;
            Xóa.Enabled = true;
        }
        private void Get_All_ChungTu()
        {
            cmd = new SqlCommand("CRUD", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@SO_CT", "");
            cmd.Parameters.AddWithValue("@TEN_NGUOI_GUI", "");
            cmd.Parameters.AddWithValue("@DIA_CHI", "");
            cmd.Parameters.AddWithValue("@SO_HS_HCC", "");
            cmd.Parameters.AddWithValue("@NGAY_HEN",DateTime.Today);
            cmd.Parameters.AddWithValue("@NGAY_NHAN", DateTime.Today);
            cmd.Parameters.AddWithValue("@TINH_NHAN", "");
            cmd.Parameters.AddWithValue("@HUYEN_NHAN", "");
            cmd.Parameters.AddWithValue("@XA_NHAN", "");
            cmd.Parameters.AddWithValue("@CUOC", "");
            cmd.Parameters.AddWithValue("@SO_HS_KEM", "0");
            cmd.Parameters.AddWithValue("@DIEN_THOAI", "");
            cmd.Parameters.AddWithValue("@MA_LOAI_HS", "");
            cmd.Parameters.AddWithValue("@TRONG_LUONG", "");
            cmd.Parameters.AddWithValue("@MA_BUUGUI", "");
            cmd.Parameters.AddWithValue("@GHI_CHU", "");
            cmd.Parameters.AddWithValue("@OperationType", "9");
            da = new SqlDataAdapter(cmd);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;

        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Them_Click(object sender, EventArgs e)
        {
            Them.Enabled = true;
            try
            {

                if (so_ct_textbox.Text != string.Empty /*&& ten_box.Text != string.Empty && txtempsalary.Text != string.Empty*/)
                {
                    cmd = new SqlCommand("CRUD", cn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@SO_CT", so_ct_textbox.Text);
                    cmd.Parameters.AddWithValue("@TEN_NGUOI_GUI", ten_box.Text);
                    cmd.Parameters.AddWithValue("@DIA_CHI", diachi_box.Text);
                    cmd.Parameters.AddWithValue("@SO_HS_HCC", so_hcc_box.Text);
                    cmd.Parameters.AddWithValue("@NGAY_HEN", dateTimePicker2.Value);
                    cmd.Parameters.AddWithValue("@NGAY_NHAN", dateTimePicker1.Value);
                    cmd.Parameters.AddWithValue("@TINH_NHAN", tinh_comboBox2.SelectedValue);
                    cmd.Parameters.AddWithValue("@HUYEN_NHAN", huyen_comboBox3.SelectedValue);
                    cmd.Parameters.AddWithValue("@XA_NHAN", xa_comboBox4.SelectedValue);
                    cmd.Parameters.AddWithValue("@CUOC", cuoc_textBox8.Text);
                    cmd.Parameters.AddWithValue("@SO_HS_KEM", sohskem.Text);
                    cmd.Parameters.AddWithValue("@DIEN_THOAI", sdt_textBox2.Text);
                    cmd.Parameters.AddWithValue("@MA_LOAI_HS", loai_hs_comboBox.SelectedValue);
                    cmd.Parameters.AddWithValue("@TRONG_LUONG", trongluong.Text);
                    cmd.Parameters.AddWithValue("@MA_BUUGUI", mabuugui.Text);
                    cmd.Parameters.AddWithValue("@GHI_CHU", ghichu_textBox9.Text);
                    cmd.Parameters.AddWithValue("@OperationType", "1");
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Bản ghi được thêm thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Get_All_ChungTu();
                    so_ct_textbox.Text ="" ;
                    ten_box.Text = "";
                    diachi_box.Text = "";
                    so_hcc_box.Text = "";
                    dateTimePicker2.Text = "";
                    dateTimePicker1.Text = "";
                    tinh_comboBox2.Text = "";
                    huyen_comboBox3.Text = "";
                    xa_comboBox4.Text = "";
                    cuoc_textBox8.Text = "";
                    sohskem.Text = "";
                    sdt_textBox2.Text = "";
                    loai_hs_comboBox.Text = "";
                    trongluong.Text = "";
                    mabuugui.Text = "";
                    ghichu_textBox9.Text = "";
                }
                else
                {
                    MessageBox.Show("Điền thiếu dữ liệu", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            
            catch
            { MessageBox.Show("Lỗi xử lý dữ liệu hoặc dữ liệu đã tồn tại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        }
        private void load_tinh()
        {
            cmd = new SqlCommand("CRUD", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@SO_CT", "");
            cmd.Parameters.AddWithValue("@TEN_NGUOI_GUI", "");
            cmd.Parameters.AddWithValue("@DIA_CHI", "");
            cmd.Parameters.AddWithValue("@SO_HS_HCC", "");
            cmd.Parameters.AddWithValue("@NGAY_HEN", DateTime.Today);
            cmd.Parameters.AddWithValue("@NGAY_NHAN", DateTime.Today);
            cmd.Parameters.AddWithValue("@TINH_NHAN", "");
            cmd.Parameters.AddWithValue("@HUYEN_NHAN", "");
            cmd.Parameters.AddWithValue("@XA_NHAN", "");
            cmd.Parameters.AddWithValue("@CUOC", "");
            cmd.Parameters.AddWithValue("@SO_HS_KEM", "0");
            cmd.Parameters.AddWithValue("@DIEN_THOAI", "");
            cmd.Parameters.AddWithValue("@MA_LOAI_HS", "");
            cmd.Parameters.AddWithValue("@TRONG_LUONG", "");
            cmd.Parameters.AddWithValue("@MA_BUUGUI", "");
            cmd.Parameters.AddWithValue("@GHI_CHU", "");
            cmd.Parameters.AddWithValue("@OperationType", "5");
            da = new SqlDataAdapter(cmd);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            tinh_comboBox5.DataSource = dt;
            tinh_comboBox5.DisplayMember = "TEN_TINH";
            tinh_comboBox5.ValueMember = "MA_TINH";
            tinh_comboBox2.DataSource = dt;
            tinh_comboBox2.DisplayMember = "TEN_TINH";
            tinh_comboBox2.ValueMember = "MA_TINH";
        }
        private void load_loai_hoso()
        {
            cmd = new SqlCommand("CRUD", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@SO_CT", "");
            cmd.Parameters.AddWithValue("@TEN_NGUOI_GUI", "");
            cmd.Parameters.AddWithValue("@DIA_CHI", "");
            cmd.Parameters.AddWithValue("@SO_HS_HCC", "");
            cmd.Parameters.AddWithValue("@NGAY_HEN", DateTime.Today);
            cmd.Parameters.AddWithValue("@NGAY_NHAN", DateTime.Today);
            cmd.Parameters.AddWithValue("@TINH_NHAN", "");
            cmd.Parameters.AddWithValue("@HUYEN_NHAN", "");
            cmd.Parameters.AddWithValue("@XA_NHAN", "");
            cmd.Parameters.AddWithValue("@CUOC", "");
            cmd.Parameters.AddWithValue("@SO_HS_KEM", "0");
            cmd.Parameters.AddWithValue("@DIEN_THOAI", "");
            cmd.Parameters.AddWithValue("@MA_LOAI_HS", "");
            cmd.Parameters.AddWithValue("@TRONG_LUONG", "");
            cmd.Parameters.AddWithValue("@MA_BUUGUI", "");
            cmd.Parameters.AddWithValue("@GHI_CHU", "");
            cmd.Parameters.AddWithValue("@OperationType", "8");
            da = new SqlDataAdapter(cmd);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            loai_hs_comboBox.DataSource = dt;
            loai_hs_comboBox.DisplayMember = "TEN_LOAI_HS";
            loai_hs_comboBox.ValueMember = "MA_LOAI_HS";
        }
        void LoadHuyen(string MaTinh)
        {
            string sql = @"select * from HUYEN_TP where MA_TINH = @MaTinh";
            SqlConnection con = cn = new SqlConnection(@"Data Source=KTNV-TIEN\SQLEXPRESS;Initial Catalog=ChungTuHCC;Integrated Security=True");
            con.Open();
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlParameter para = new SqlParameter("@MaTinh", SqlDbType.NVarChar);
            para.Value = MaTinh;
            cmd.Parameters.Add(para);
            da = new SqlDataAdapter(cmd);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
                huyen_comboBox3.DataSource = dt;
                huyen_comboBox3.DisplayMember = "TEN_HUYENTP";
                huyen_comboBox3.ValueMember = "MA_HUYENTP";
                huyen_comboBox6.DataSource = dt;
            huyen_comboBox6.DisplayMember = "TEN_HUYENTP";
            huyen_comboBox6.ValueMember = "MA_HUYENTP";
        }
        void LoadXa(string MaHuyen)
        {
            string sql = @"select * from XAPHUONG where MA_HUYENTP= @MaHuyen";
            SqlConnection con = cn = new SqlConnection(@"Data Source=KTNV-TIEN\SQLEXPRESS;Initial Catalog=ChungTuHCC;Integrated Security=True");
            con.Open();
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlParameter para = new SqlParameter("@MaHuyen", SqlDbType.NVarChar);
            para.Value = MaHuyen;
            cmd.Parameters.Add(para);
            da = new SqlDataAdapter(cmd);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            xa_comboBox4.DataSource = dt;
            xa_comboBox4.DisplayMember = "TEN_XA_PHUONG";
            xa_comboBox4.ValueMember = "MA_XA_PHUONG";
        }
        private void tinh_comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadHuyen(tinh_comboBox5.SelectedValue.ToString());
        }

        private void tinh_comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadHuyen(tinh_comboBox2.SelectedValue.ToString());
        }
        private void huyen_comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {   
            LoadXa(huyen_comboBox3.SelectedValue.ToString());
        }

        private void huyen_comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadXa(huyen_comboBox3.SelectedValue.ToString());
        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void Xóa_Click(object sender, EventArgs e)
        {
            cmd = new SqlCommand("CRUD", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@SO_CT", so_ct_textbox.Text);
            cmd.Parameters.AddWithValue("@TEN_NGUOI_GUI", ten_box.Text);
            cmd.Parameters.AddWithValue("@DIA_CHI", diachi_box.Text);
            cmd.Parameters.AddWithValue("@SO_HS_HCC", so_hcc_box.Text);
            cmd.Parameters.AddWithValue("@NGAY_HEN", dateTimePicker2.Value);
            cmd.Parameters.AddWithValue("@NGAY_NHAN", dateTimePicker1.Value);
            cmd.Parameters.AddWithValue("@TINH_NHAN", tinh_comboBox2.SelectedValue);
            cmd.Parameters.AddWithValue("@HUYEN_NHAN", huyen_comboBox3.SelectedValue);
            cmd.Parameters.AddWithValue("@XA_NHAN", xa_comboBox4.SelectedValue);
            cmd.Parameters.AddWithValue("@CUOC", cuoc_textBox8.Text);
            cmd.Parameters.AddWithValue("@SO_HS_KEM", sohskem.Text);
            cmd.Parameters.AddWithValue("@DIEN_THOAI", sdt_textBox2.Text);
            cmd.Parameters.AddWithValue("@MA_LOAI_HS", loai_hs_comboBox.SelectedValue);
            cmd.Parameters.AddWithValue("@TRONG_LUONG", trongluong.Text);
            cmd.Parameters.AddWithValue("@MA_BUUGUI", mabuugui.Text);
            cmd.Parameters.AddWithValue("@GHI_CHU", ghichu_textBox9.Text);
            cmd.Parameters.AddWithValue("@OperationType", "3");
            cmd.ExecuteNonQuery();
            MessageBox.Show("Bản ghi được xóa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Get_All_ChungTu();
            so_ct_textbox.Text = "";
            ten_box.Text = "";
            diachi_box.Text = "";
            so_hcc_box.Text = "";
            dateTimePicker2.Text = "";
            dateTimePicker1.Text = "";
            tinh_comboBox2.Text = "";
            huyen_comboBox3.Text = "";
            xa_comboBox4.Text = "";
            cuoc_textBox8.Text = "";
            sohskem.Text = "";
            sdt_textBox2.Text = "";
            loai_hs_comboBox.Text = "";
            trongluong.Text = "";
            mabuugui.Text = "";
            ghichu_textBox9.Text = "";
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            so_ct_textbox.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            ten_box.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            diachi_box.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            so_hcc_box.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            dateTimePicker2.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            dateTimePicker1.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            tinh_comboBox2.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            huyen_comboBox3.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            xa_comboBox4.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            cuoc_textBox8.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
            sohskem.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
            sdt_textBox2.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
            loai_hs_comboBox.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
            trongluong.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
            mabuugui.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
            ghichu_textBox9.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();
            //so_ct_textbox.Enabled = false;    
        }
        private void Sửa_Click(object sender, EventArgs e)
        {
            if (so_ct_textbox.Text != string.Empty )
            {
                cmd = new SqlCommand("CRUD", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@SO_CT", so_ct_textbox.Text);
                cmd.Parameters.AddWithValue("@TEN_NGUOI_GUI", ten_box.Text);
                cmd.Parameters.AddWithValue("@DIA_CHI", diachi_box.Text);
                cmd.Parameters.AddWithValue("@SO_HS_HCC", so_hcc_box.Text);
                cmd.Parameters.AddWithValue("@NGAY_HEN", dateTimePicker2.Value);
                cmd.Parameters.AddWithValue("@NGAY_NHAN", dateTimePicker1.Value);
                cmd.Parameters.AddWithValue("@TINH_NHAN", tinh_comboBox2.SelectedValue);
                cmd.Parameters.AddWithValue("@HUYEN_NHAN", huyen_comboBox3.SelectedValue);
                cmd.Parameters.AddWithValue("@XA_NHAN", xa_comboBox4.SelectedValue);
                cmd.Parameters.AddWithValue("@CUOC", cuoc_textBox8.Text);
                cmd.Parameters.AddWithValue("@SO_HS_KEM", sohskem.Text);
                cmd.Parameters.AddWithValue("@DIEN_THOAI", sdt_textBox2.Text);
                cmd.Parameters.AddWithValue("@MA_LOAI_HS", loai_hs_comboBox.SelectedValue);
                cmd.Parameters.AddWithValue("@TRONG_LUONG", trongluong.Text);
                cmd.Parameters.AddWithValue("@MA_BUUGUI", mabuugui.Text);
                cmd.Parameters.AddWithValue("@GHI_CHU", ghichu_textBox9.Text);
                cmd.Parameters.AddWithValue("@OperationType", "2");
                cmd.ExecuteNonQuery();
                MessageBox.Show("Bản ghi được sửa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Get_All_ChungTu();
                so_ct_textbox.Text = "";
                ten_box.Text = "";
                diachi_box.Text = "";
                so_hcc_box.Text = "";
                dateTimePicker2.Text = "";
                dateTimePicker1.Text = "";
                tinh_comboBox2.Text = "";
                huyen_comboBox3.Text = "";
                xa_comboBox4.Text = "";
                cuoc_textBox8.Text = "";
                sohskem.Text = "";
                sdt_textBox2.Text = "";
                loai_hs_comboBox.Text = "";
                trongluong.Text = "";
                mabuugui.Text = "";
                ghichu_textBox9.Text = "";
                //Xóa.Enabled = false;
              
            }
            else
            {
                MessageBox.Show("Điền thiếu dữ liệu hoặc số chứng từ đã tồn tại", "Invalid Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void xuatExcel_Click(object sender, EventArgs e)
        {
            // creating Excel Application  
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application  
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook  
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            // see the excel sheet behind the program  
            app.Visible = true;
            // get the reference of first sheet. By default its name is Sheet1.  
            // store its reference to worksheet  
            worksheet = (_Worksheet)workbook.Sheets["Sheet1"];
            worksheet = (_Worksheet)workbook.ActiveSheet;
            // changing the name of active sheet  
            worksheet.Name = "Exported from gridview";
            // lấy dữ liệu tên cột
            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }
            // đưa dữ liệu từ datagridview ra worksheet
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }
            // save the application  
            //workbook.Save("D:\\baocaoexcel.xls", XlSaveAsAccessMode.xlExclusive);
            //System.Windows.Forms.SaveFileDialog saveDlg = new System.Windows.Forms.SaveFileDialog();
            //saveDlg.InitialDirectory = @"D:\";
            //saveDlg.Filter = "Excel files (*.xls)|*.xlsx";
            //saveDlg.FilterIndex = 0;    
            //saveDlg.RestoreDirectory = true;
            //saveDlg.Title = "Export Excel File To";
            //if (saveDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{
            //    string path = saveDlg.FileName;
            //    workbook.SaveCopyAs(path);
            //    workbook.Saved = true;
            //    workbook.Close(true);
            //    //app.Quit();
            //}
        }
    }
}
