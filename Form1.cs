using System;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Spire.Doc;
using Spire.DataExport;
using Spire.DataExport.RTF;
using System.Drawing;
using DocumentFormat.OpenXml.Drawing;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
using System.IO;
using System.Data.OleDb;
using System.ComponentModel;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Globalization;
using ClosedXML.Excel;
namespace XuatExcelApp
{
    public partial class Form1 : Form
    {
        SqlConnection cn = new SqlConnection(@"Data Source=DEVLAP\SQLEXPRESS;Initial Catalog=ChungTuHCC;Integrated Security=True");
        SqlCommand cmd;
        SqlDataAdapter da;
        SqlDataReader dr;
        public static int kieu_bao_cao = 2;
        public static string day;
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            cn.Open();
            //bind data in data grid view  
            Get_All_ChungTu();
            load_tinh();
            load_tinh_tt();
            load_loai_hoso();
            int xid;
            bool parseOK1 = Int32.TryParse(tinh_comboBox2.SelectedValue.ToString(), out xid);
            LoadHuyen(xid);
            int zid;
            bool parseOK2 = Int32.TryParse(tinh_comboBox5.SelectedValue.ToString(), out zid);
            LoadHuyen_tt(zid);
            int fid;
            bool parseOK = Int32.TryParse(huyen_comboBox3.SelectedValue.ToString(), out fid);
            LoadXa(fid);
            ////disable delete and update button on load  
            Sửa.Enabled = true;
            Xóa.Enabled = true;
            //auto_add_id();
            xoa_noidung_box();
            cn.Close();
        }
        private void Get_All_ChungTu()
        {
            cmd = new SqlCommand("CRUD", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@SO_CT", 0);
            cmd.Parameters.AddWithValue("@Approved", 0);
            cmd.Parameters.AddWithValue("@TEN_NGUOI_NHAN", "");
            cmd.Parameters.AddWithValue("@DIA_CHI", "");
            cmd.Parameters.AddWithValue("@SO_HS_HCC", "");
            cmd.Parameters.AddWithValue("@NGAY_HEN", "");
            cmd.Parameters.AddWithValue("@NGAY_NHAN", "");
            cmd.Parameters.AddWithValue("@DIA_CHI_THUONG_TRU", "");
            cmd.Parameters.AddWithValue("@TINH_NHAN", "");
            cmd.Parameters.AddWithValue("@HUYEN_NHAN", "");
            cmd.Parameters.AddWithValue("@XA_NHAN", "");
            cmd.Parameters.AddWithValue("@SO_HS_KEM", "0");
            cmd.Parameters.AddWithValue("@DIEN_THOAI", "");
            cmd.Parameters.AddWithValue("@MA_LOAI_HS", "");
            cmd.Parameters.AddWithValue("@TRONG_LUONG", "");
            cmd.Parameters.AddWithValue("@MA_BUUGUI", "");
            cmd.Parameters.AddWithValue("@GHI_CHU", "");
            cmd.Parameters.AddWithValue("@NGAY_DUYET", "");
            cmd.Parameters.AddWithValue("@CUOC_PHI", decimal.Zero);
            cmd.Parameters.AddWithValue("@NGAY_NHAP", "");
            cmd.Parameters.AddWithValue("@OperationType", "9");
            da = new SqlDataAdapter(cmd);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }
        private void search_box_tinh()
           { 
            tinh_comboBox2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            tinh_comboBox2.AutoCompleteSource = AutoCompleteSource.ListItems;
            cmd = new SqlCommand("search_box_tinh", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            da = new SqlDataAdapter(cmd);
            cmd.Parameters.AddWithValue("@box", tinh_comboBox2.Text.ToString());
            System.Data.DataTable tinh = new System.Data.DataTable();
            DataView dtview = new DataView(tinh);
            da.Fill(tinh);
            dtview.Sort = "TEN_TINH DESC";
            comboBox1.DataSource = tinh;
            comboBox1.ValueMember = "TEN_TINH";
            comboBox1.DisplayMember = "TEN_TINH";
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
            if(Them.Enabled==false)
            {
                Them.Enabled = true;
                xoa_noidung_box();
                auto_add_id();
            }    
            if (DialogResult.Yes == MessageBox.Show("Xác nhận trước khi thêm", "Cảnh báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                if (cn.State == ConnectionState.Closed)
                {
                    cn.Open();
                }
                Them.Enabled = true;
                try
                {
                    if (comboBox1.Text != string.Empty /*&& ten_box.Text != string.Empty && txtempsalary.Text != string.Empty*/)
                    {
                        cmd = new SqlCommand("CRUD", cn);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@SO_CT", int.Parse(comboBox1.Text));
                        cmd.Parameters.AddWithValue("@Approved", 0);
                        cmd.Parameters.AddWithValue("@TEN_NGUOI_NHAN", ten_box.Text);
                        cmd.Parameters.AddWithValue("@DIA_CHI", diachi_box.Text);
                        try
                        {
                            int sohcc = int.Parse(so_hcc_box.Text.ToString());
                            cmd.Parameters.AddWithValue("@SO_HS_HCC", so_hcc_box.Text);
                        }
                        catch
                        {
                            so_hcc_box.Text = "";
                            MessageBox.Show("Số HCC không hợp lệ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        cmd.Parameters.AddWithValue("@NGAY_HEN", dateTimePicker2.Value.ToShortDateString());
                        cmd.Parameters.AddWithValue("@NGAY_NHAN", dateTimePicker1.Value.Date.ToShortDateString());
                        cmd.Parameters.AddWithValue("@DIA_CHI_THUONG_TRU", (diachi_thuongtru_box.Text + " " + tinh_comboBox5.Text + " " + huyen_comboBox6.Text).ToString());
                        try 
                        {   
                            cmd.Parameters.AddWithValue("@TINH_NHAN", int.Parse(tinh_comboBox2.SelectedValue.ToString()));
                            cmd.Parameters.AddWithValue("@HUYEN_NHAN", int.Parse(huyen_comboBox3.SelectedValue.ToString()));
                            cmd.Parameters.AddWithValue("@XA_NHAN", int.Parse(xa_comboBox4.SelectedValue.ToString()));
                        }
                        catch
                        {
                            MessageBox.Show("Dữ liệu tỉnh, huyện, xã không hợp lệ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                       
                        try
                        {
                            int shskem = int.Parse(sohskem.Text.ToString());
                            cmd.Parameters.AddWithValue("@SO_HS_KEM", sohskem.Text);
                        }
                        catch
                        {
                            sohskem.Text = "";
                            MessageBox.Show("Số HS kèm không hợp lệ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        try
                        {   
                            int sdt = int.Parse(sdt_textBox2.Text.ToString());
                            cmd.Parameters.AddWithValue("@DIEN_THOAI", sdt_textBox2.Text);
                        }
                        catch
                        {
                            sdt_textBox2.Text = "";
                            MessageBox.Show("Số điện thoại không hợp lệ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        cmd.Parameters.AddWithValue("@MA_LOAI_HS", int.Parse(loai_hs_comboBox.SelectedValue.ToString()));
                        cmd.Parameters.AddWithValue("@TRONG_LUONG", trongluong.Text);
                        cmd.Parameters.AddWithValue("@MA_BUUGUI", mabuugui.Text);
                        cmd.Parameters.AddWithValue("@GHI_CHU", ghichu_textBox9.Text);
                        cmd.Parameters.AddWithValue("@NGAY_DUYET", "");
                        cmd.Parameters.AddWithValue("@CUOC_PHI", Decimal.Parse(cuoc_textBox8.Text.ToString()));
                        cmd.Parameters.AddWithValue("@NGAY_NHAP", DateTime.Now.ToShortDateString());
                        cmd.Parameters.AddWithValue("@OperationType", "1");
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Bản ghi được thêm thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Get_All_ChungTu();

                    }
                    else
                    {
                        MessageBox.Show("Điền thiếu dữ liệu", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    xoa_noidung_box();
                }
                catch
                {
                    MessageBox.Show("Số CT bị trùng hoặc dữ liệu không hợp lệ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                }
                xoa_noidung_box();
                cn.Close();
            }
        }
        private void xoa_noidung_box()
        {
            //auto_add_id();
            dateTimePicker1.Text = "";
            dateTimePicker2.Text = "";
            loai_hs_comboBox.Text = "";
            dateTimePicker3.Text = "";
            ten_box.Text = "";
            diachi_box.Text = "";
            so_hcc_box.Text = "";
            dateTimePicker4.Text = "";
            diachi_thuongtru_box.Text = "";
            tinh_comboBox5.Text = "";
            tinh_comboBox2.Text = "";
            huyen_comboBox3.Text = "";
            huyen_comboBox6.Text = "";
            xa_comboBox4.Text = "";
            sohskem.Text = "";
            trongluong.Text = "";
            mabuugui.Text = "";
            cuoc_textBox8.Text = "";
            ghichu_textBox9.Text = "";
            sdt_textBox2.Text = "";
        }
        private void load_tinh()
        {
            cmd = new SqlCommand("CRUD", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@SO_CT", 0);
            cmd.Parameters.AddWithValue("@Approved", 0);
            cmd.Parameters.AddWithValue("@TEN_NGUOI_NHAN", "");
            cmd.Parameters.AddWithValue("@DIA_CHI", "");
            cmd.Parameters.AddWithValue("@SO_HS_HCC", "");
            cmd.Parameters.AddWithValue("@NGAY_HEN", "");
            cmd.Parameters.AddWithValue("@NGAY_NHAN", "");
            cmd.Parameters.AddWithValue("@DIA_CHI_THUONG_TRU", "");
            cmd.Parameters.AddWithValue("@TINH_NHAN", "");
            cmd.Parameters.AddWithValue("@HUYEN_NHAN", "");
            cmd.Parameters.AddWithValue("@XA_NHAN", "");
            cmd.Parameters.AddWithValue("@SO_HS_KEM", "0");
            cmd.Parameters.AddWithValue("@DIEN_THOAI", "");
            cmd.Parameters.AddWithValue("@MA_LOAI_HS", "");
            cmd.Parameters.AddWithValue("@TRONG_LUONG", "");
            cmd.Parameters.AddWithValue("@MA_BUUGUI", "");
            cmd.Parameters.AddWithValue("@GHI_CHU", "");
            cmd.Parameters.AddWithValue("@NGAY_DUYET", "");
            cmd.Parameters.AddWithValue("@CUOC_PHI", decimal.Zero);
            cmd.Parameters.AddWithValue("@NGAY_NHAP", "");
            cmd.Parameters.AddWithValue("@OperationType", "5");
            da = new SqlDataAdapter(cmd);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            tinh_comboBox2.DataSource = dt;
            tinh_comboBox2.DisplayMember = "TEN_TINH";
            tinh_comboBox2.ValueMember = "MA_TINH";
        }
        private void load_tinh_tt()
        {
            cmd = new SqlCommand("CRUD", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@SO_CT", 0);
            cmd.Parameters.AddWithValue("@Approved", 0);
            cmd.Parameters.AddWithValue("@TEN_NGUOI_NHAN", "");
            cmd.Parameters.AddWithValue("@DIA_CHI", "");
            cmd.Parameters.AddWithValue("@SO_HS_HCC", "");
            cmd.Parameters.AddWithValue("@NGAY_HEN", "");
            cmd.Parameters.AddWithValue("@NGAY_NHAN", "");
            cmd.Parameters.AddWithValue("@DIA_CHI_THUONG_TRU", "");
            cmd.Parameters.AddWithValue("@TINH_NHAN", "");
            cmd.Parameters.AddWithValue("@HUYEN_NHAN", "");
            cmd.Parameters.AddWithValue("@XA_NHAN", "");
            cmd.Parameters.AddWithValue("@SO_HS_KEM", "0");
            cmd.Parameters.AddWithValue("@DIEN_THOAI", "");
            cmd.Parameters.AddWithValue("@MA_LOAI_HS", "");
            cmd.Parameters.AddWithValue("@TRONG_LUONG", "");
            cmd.Parameters.AddWithValue("@MA_BUUGUI", "");
            cmd.Parameters.AddWithValue("@GHI_CHU", "");
            cmd.Parameters.AddWithValue("@NGAY_DUYET", "1/1/1900");
            cmd.Parameters.AddWithValue("@CUOC_PHI", decimal.Zero);
            cmd.Parameters.AddWithValue("@NGAY_NHAP", "");
            cmd.Parameters.AddWithValue("@OperationType", "5");
            da = new SqlDataAdapter(cmd);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            tinh_comboBox5.DataSource = dt;
            tinh_comboBox5.DisplayMember = "TEN_TINH";
            tinh_comboBox5.ValueMember = "MA_TINH";
        }
        private void load_loai_hoso()
        {
            cmd = new SqlCommand("CRUD", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@SO_CT", 0);
            cmd.Parameters.AddWithValue("@Approved", 0);
            cmd.Parameters.AddWithValue("@TEN_NGUOI_NHAN", "");
            cmd.Parameters.AddWithValue("@DIA_CHI", "");
            cmd.Parameters.AddWithValue("@SO_HS_HCC", "");
            cmd.Parameters.AddWithValue("@NGAY_HEN", "");
            cmd.Parameters.AddWithValue("@NGAY_NHAN", "");
            cmd.Parameters.AddWithValue("@DIA_CHI_THUONG_TRU", diachi_thuongtru_box.Text);
            cmd.Parameters.AddWithValue("@TINH_NHAN", "");
            cmd.Parameters.AddWithValue("@HUYEN_NHAN", "");
            cmd.Parameters.AddWithValue("@XA_NHAN", "");
            cmd.Parameters.AddWithValue("@SO_HS_KEM", "0");
            cmd.Parameters.AddWithValue("@DIEN_THOAI", "");
            cmd.Parameters.AddWithValue("@MA_LOAI_HS", "");
            cmd.Parameters.AddWithValue("@TRONG_LUONG", "");
            cmd.Parameters.AddWithValue("@MA_BUUGUI", "");
            cmd.Parameters.AddWithValue("@GHI_CHU", "");
            cmd.Parameters.AddWithValue("@NGAY_DUYET", "");
            cmd.Parameters.AddWithValue("@CUOC_PHI", decimal.Zero);
            cmd.Parameters.AddWithValue("@NGAY_NHAP", "");
            cmd.Parameters.AddWithValue("@OperationType", "8");
            da = new SqlDataAdapter(cmd);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            loai_hs_comboBox.DataSource = dt;
            loai_hs_comboBox.DisplayMember = "TEN_LOAI_HS";
            loai_hs_comboBox.ValueMember = "MA_LOAI_HS";
        }
        void LoadHuyen(int MaTinh)
        {
            SqlCommand cmd = new SqlCommand("load_huyen", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add(new SqlParameter("@tinh", MaTinh));
            da = new SqlDataAdapter(cmd);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            huyen_comboBox3.DataSource = dt;
            huyen_comboBox3.DisplayMember = "TEN_HUYENTP";
            huyen_comboBox3.ValueMember = "MA_HUYENTP";
            //huyen_comboBox6.DataSource = dt;
            //huyen_comboBox6.DisplayMember = "TEN_HUYENTP";
            //huyen_comboBox6.ValueMember = "MA_HUYENTP";
        }
        void LoadHuyen_tt(int MaTinh)
        {
            SqlCommand cmd = new SqlCommand("load_huyen", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add(new SqlParameter("@tinh", MaTinh));
            da = new SqlDataAdapter(cmd);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            //huyen_comboBox3.DataSource = dt;
            //huyen_comboBox3.DisplayMember = "TEN_HUYENTP";
            //huyen_comboBox3.ValueMember = "MA_HUYENTP";
            huyen_comboBox6.DataSource = dt;
            huyen_comboBox6.DisplayMember = "TEN_HUYENTP";
            huyen_comboBox6.ValueMember = "MA_HUYENTP";
        }
        void LoadXa(int MaHuyen)
        {
            SqlCommand cmd = new SqlCommand("load_xa", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add(new SqlParameter("@huyen", MaHuyen));
            da = new SqlDataAdapter(cmd);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            xa_comboBox4.DataSource = dt;
            xa_comboBox4.DisplayMember = "TEN_XA_PHUONG";
            xa_comboBox4.ValueMember = "MA_XA_PHUONG";
        }
        private void label20_Click(object sender, EventArgs e)
        {
        }
        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }
        private void Xóa_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("Xác nhận trước khi xóa", "Cảnh báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                if (cn.State == ConnectionState.Closed)
                {
                    cn.Open();
                }
                try
                {
                    cmd = new SqlCommand("CRUD", cn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@SO_CT", int.Parse(comboBox1.Text));
                    cmd.Parameters.AddWithValue("@Approved", 0);
                    cmd.Parameters.AddWithValue("@TEN_NGUOI_NHAN", ten_box.Text);
                    cmd.Parameters.AddWithValue("@DIA_CHI", diachi_box.Text);
                    cmd.Parameters.AddWithValue("@SO_HS_HCC", so_hcc_box.Text);
                    cmd.Parameters.AddWithValue("@NGAY_HEN", dateTimePicker2.Value.ToShortDateString());
                    cmd.Parameters.AddWithValue("@NGAY_NHAN", dateTimePicker1.Value.ToShortDateString());
                    cmd.Parameters.AddWithValue("@DIA_CHI_THUONG_TRU", (diachi_thuongtru_box.Text + ", " + tinh_comboBox5.Text + ", " + huyen_comboBox6.Text).ToString());
                    cmd.Parameters.AddWithValue("@TINH_NHAN", int.Parse(tinh_comboBox2.SelectedValue.ToString()));
                    cmd.Parameters.AddWithValue("@HUYEN_NHAN", int.Parse(huyen_comboBox3.SelectedValue.ToString()));
                    cmd.Parameters.AddWithValue("@XA_NHAN", int.Parse(xa_comboBox4.SelectedValue.ToString()));
                    cmd.Parameters.AddWithValue("@SO_HS_KEM", sohskem.Text);
                    cmd.Parameters.AddWithValue("@DIEN_THOAI", sdt_textBox2.Text);
                    cmd.Parameters.AddWithValue("@MA_LOAI_HS", int.Parse(loai_hs_comboBox.SelectedValue.ToString()));
                    cmd.Parameters.AddWithValue("@TRONG_LUONG", trongluong.Text);
                    cmd.Parameters.AddWithValue("@MA_BUUGUI", mabuugui.Text);
                    cmd.Parameters.AddWithValue("@GHI_CHU", ghichu_textBox9.Text);
                    cmd.Parameters.AddWithValue("@NGAY_DUYET", dateTimePicker3.Value.ToShortDateString());
                    cmd.Parameters.AddWithValue("@CUOC_PHI", Decimal.Parse(cuoc_textBox8.Text.ToString()));
                    cmd.Parameters.AddWithValue("@NGAY_NHAP", "");
                    cmd.Parameters.AddWithValue("@OperationType", "3");
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Bản ghi được xóa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Get_All_ChungTu();
                }
                catch
                {
                    MessageBox.Show("Số CT bị trùng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                cn.Close();
            }
            Them.Enabled = true;
            xoa_noidung_box();
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (cn.State != ConnectionState.Open)
            {
                cn.Open();
            }
            try
            {
            Them.Enabled = false;
            comboBox1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                if (Convert.ToBoolean(dataGridView1.CurrentRow.Cells[1].EditedFormattedValue) == true)
            {
                duyet_box.CheckState = CheckState.Checked;
            }
            else
            {
                duyet_box.CheckState = CheckState.Unchecked;
            }
            dataGridView1.CurrentRow.Cells[1].ReadOnly = true;
            ten_box.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            diachi_box.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            so_hcc_box.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            //DateTime date1 = DateTime.ParseExact(dataGridView1.CurrentRow.Cells[5].Value.ToString(), "dd'|'MM'|'yyyy", null);
            dateTimePicker2.Text = (DateTime.ParseExact(dataGridView1.CurrentRow.Cells[5].Value.ToString(), "dd'/'MM'/'yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None)).ToString();
            ////DateTime date2 = DateTime.ParseExact(dataGridView1.CurrentRow.Cells[6].Value.ToString(), "dd'|'MM'|'yyyy", System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None);
            dateTimePicker1.Text = (DateTime.ParseExact(dataGridView1.CurrentRow.Cells[6].Value.ToString(), "dd'/'MM'/'yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None)).ToString();
            diachi_thuongtru_box.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            SqlCommand cmd = new SqlCommand("load_tinh_huyen_xa", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add(new SqlParameter("@tinh", int.Parse(dataGridView1.CurrentRow.Cells[8].Value.ToString())));
            cmd.Parameters.Add(new SqlParameter("@huyen", int.Parse(dataGridView1.CurrentRow.Cells[9].Value.ToString())));
            cmd.Parameters.Add(new SqlParameter("@xa", int.Parse(dataGridView1.CurrentRow.Cells[10].Value.ToString())));
            da = new SqlDataAdapter(cmd);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            tinh_comboBox2.Text = dt.Rows[0][0].ToString();
            huyen_comboBox3.Text = dt.Rows[0][1].ToString();
            xa_comboBox4.Text = dt.Rows[0][2].ToString();
            sohskem.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
            sdt_textBox2.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
            SqlCommand cmd1 = new SqlCommand("load_loai_hoso", cn);
            cmd1.CommandType = CommandType.StoredProcedure;
            cmd1.Parameters.Add(new SqlParameter("@id", int.Parse(dataGridView1.CurrentRow.Cells[13].Value.ToString())));
            SqlDataAdapter da1 = new SqlDataAdapter(cmd1);
            System.Data.DataTable dt1 = new System.Data.DataTable();
            da1.Fill(dt1);
            loai_hs_comboBox.Text = dt1.Rows[0][0].ToString();
            trongluong.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
            mabuugui.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();
            ghichu_textBox9.Text = dataGridView1.CurrentRow.Cells[16].Value.ToString();
            if (dataGridView1.CurrentRow.Cells[17].Value.ToString() != "")
            {
                dateTimePicker3.Visible = true;
                dateTimePicker3.Text = (DateTime.ParseExact(dataGridView1.CurrentRow.Cells[17].Value.ToString(), "dd'/'MM'/'yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None)).ToString();
            }
            if (dataGridView1.CurrentRow.Cells[17].Value.ToString() == "")
            {
                dateTimePicker3.Visible = false;
            }
            cuoc_textBox8.Text = dataGridView1.CurrentRow.Cells[18].Value.ToString();
            }
            catch
            {
                loai_hs_comboBox.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
                MessageBox.Show("Hàng trống dữ liệu", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                xoa_noidung_box();
            }
        }
        private void Sửa_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("Xác nhận trước khi sửa", "Cảnh báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                if (cn.State == ConnectionState.Closed)
                {
                    cn.Open();
                }
                try
                {
                    if (comboBox1.Text != string.Empty)
                    {
                        int i;
                        if (duyet_box.Checked == true)
                        {
                            i = 1;
                        }
                        else i = 0;
                        cmd = new SqlCommand("CRUD", cn);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@SO_CT", int.Parse(comboBox1.Text));
                        cmd.Parameters.AddWithValue("@Approved", i);
                        cmd.Parameters.AddWithValue("@TEN_NGUOI_NHAN", ten_box.Text);
                        cmd.Parameters.AddWithValue("@DIA_CHI", diachi_box.Text);
                        
                        try
                        {
                            int sohcc = int.Parse(so_hcc_box.Text.ToString());
                            cmd.Parameters.AddWithValue("@SO_HS_HCC", so_hcc_box.Text);
                        }
                        catch
                        {
                            so_hcc_box.Text = "";
                            MessageBox.Show("Số HCC không hợp lệ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        cmd.Parameters.AddWithValue("@NGAY_HEN", dateTimePicker2.Value.Date.ToShortDateString());
                        cmd.Parameters.AddWithValue("@NGAY_NHAN", dateTimePicker1.Value.Date.ToShortDateString());
                        if (!tinh_comboBox5.SelectedValue.ToString().Equals("0"))
                        {
                        cmd.Parameters.AddWithValue("@DIA_CHI_THUONG_TRU", (diachi_thuongtru_box.Text + " " + tinh_comboBox5.Text + " " + huyen_comboBox6.Text).ToString());
                        }
                        if (tinh_comboBox5.SelectedValue.ToString().Equals("0"))
                        {
                            cmd.Parameters.AddWithValue("@DIA_CHI_THUONG_TRU", diachi_thuongtru_box.Text);
                        }
                        cmd.Parameters.AddWithValue("@TINH_NHAN", int.Parse(tinh_comboBox2.SelectedValue.ToString()));
                        cmd.Parameters.AddWithValue("@HUYEN_NHAN", int.Parse(huyen_comboBox3.SelectedValue.ToString()));
                        cmd.Parameters.AddWithValue("@XA_NHAN", int.Parse(xa_comboBox4.SelectedValue.ToString()));   
                        try
                        {
                            int sohs = int.Parse(sohskem.Text.ToString());
                            cmd.Parameters.AddWithValue("@SO_HS_KEM", sohskem.Text);
                        }
                        catch
                        {
                            sohskem.Text = "";
                            MessageBox.Show("Số hồ sơ kèm không hợp lệ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        try
                        {
                            int shs = int.Parse(sdt_textBox2.Text.ToString());
                            cmd.Parameters.AddWithValue("@DIEN_THOAI", sdt_textBox2.Text);
                        }
                        catch
                        {
                            sdt_textBox2.Text = "";
                            MessageBox.Show("Số điện thoại không hợp lệ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        cmd.Parameters.AddWithValue("@MA_LOAI_HS", int.Parse(loai_hs_comboBox.SelectedValue.ToString()));
                        cmd.Parameters.AddWithValue("@TRONG_LUONG", trongluong.Text);
                        cmd.Parameters.AddWithValue("@MA_BUUGUI", mabuugui.Text);
                        cmd.Parameters.AddWithValue("@GHI_CHU", ghichu_textBox9.Text);
                        if (dateTimePicker3.Visible == true)
                        {
                            cmd.Parameters.AddWithValue("@NGAY_DUYET", dateTimePicker3.Value.ToShortDateString());
                        }
                        if (dateTimePicker3.Visible == false)
                        {
                            cmd.Parameters.AddWithValue("@NGAY_DUYET", "");
                        }
                        cmd.Parameters.AddWithValue("@CUOC_PHI", Decimal.Parse(cuoc_textBox8.Text.ToString()));
                        cmd.Parameters.AddWithValue("@NGAY_NHAP", dataGridView1.CurrentRow.Cells[19].Value.ToString());
                        cmd.Parameters.AddWithValue("@OperationType", "2");
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Bản ghi được sửa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Get_All_ChungTu();
                    }
                    else
                    {
                        MessageBox.Show("Lỗi dữ liệu nhập vào hoặc số chứng từ đã tồn tại", "Invalid Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch
                { 
                    MessageBox.Show("Số CT bị trùng/Dữ liệu đầu vào không hợp lệ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                }
                cn.Close();
            }
            Them.Enabled = true;
            xoa_noidung_box();

        }
        private void xuatExcel_Click(object sender, EventArgs e)
        {
            // creating Excel Application  
            //Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            //// creating new WorkBook within Excel application  
            //Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            //// creating new Excelsheet in workbook  
            //Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            //// see the excel sheet behind the program  
            //app.Visible = true;
            //// get the reference of first sheet. By default its name is Sheet1.  
            //// store its reference to worksheet  
            //worksheet = (_Worksheet)workbook.Sheets["Sheet1"];
            //worksheet = (_Worksheet)workbook.ActiveSheet;
            //// changing the name of active sheet  
            //worksheet.Name = "Sheet1";
            //worksheet.Columns.ColumnWidth =20;
            //// lấy dữ liệu tên cột
            //for (int i = 1; i <= dataGridView1.Columns.Count; i++)
            //{
            //    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            //}
            //// đưa dữ liệu từ datagridview ra worksheet
            //for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            //{
            //    for (int j = 0; j < dataGridView1.Columns.Count; j++)
            //    {
            //        worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
            //    }
            //}
            //Adding the Columns
            int k = dataGridView1.Columns.Count;
            int l = dataGridView1.Rows.Count;
            if (dataGridView1.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.ApplicationClass XcelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
                XcelApp.Application.Workbooks.Add(Type.Missing);
                //Sheet trong excel có chỉ số bắt đầu bằng 1
                for (int i = 0; i < k; i++)
                {
                    XcelApp.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;//Gán header cho hàng đầu excel
                }
                for (int i = 2; i <= l; i++)
                {
                    for (int j = 1; j <= k; j++)
                    {
                        XcelApp.Cells[i, j] = dataGridView1.Rows[i - 2].Cells[j - 1].Value.ToString();//gán giá trị cho từ thứ 2 đến hàng cuối.
                    }
                }
                XcelApp.Columns.AutoFit();
                XcelApp.Visible = true;
            }
        }
        private void button7_Click(object sender, EventArgs e)
        {
            //IN_BAOCAO_THEO_HUYEN(int.Parse((tinh_comboBox5.SelectedValue).ToString()), int.Parse((huyen_comboBox6.SelectedValue).ToString()));
            IN_BAO_ALL(dateTimePicker4.Value.ToShortDateString());
        }
        void IN_BAOCAO_THEO_HUYEN(int ma_tinh, int ma_huyen)
        {
            if (cn.State != ConnectionState.Open)
            {
                cn.Open();
            }
            SqlCommand cmd = new SqlCommand("dem_ban_ghi", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add(new SqlParameter("@tinh", ma_tinh));
            cmd.Parameters.Add(new SqlParameter("@huyen", ma_huyen));
            //Add parameters like this
            int count = (int)cmd.ExecuteScalar();
        }
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
        }
        private void label15_Click(object sender, EventArgs e)
        {
        }
        private void label21_Click(object sender, EventArgs e)
        {
        }
        private void cuoc_textBox8_TextChanged(object sender, EventArgs e)
        {
        }
        private void printPreviewDialog1_Load(object sender, EventArgs e)
        {
        }
        private void tenhuyen_report_TextChanged(object sender, EventArgs e)
        {
        }
        private void baocao_TextChanged(object sender, EventArgs e)
        {
        }
        private void so_ct_textbox_TextChanged(object sender, EventArgs e)
        {
        }
        void auto_add_id()
        {
            SqlCommand cmd = new SqlCommand("auto_add_id", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            da = new SqlDataAdapter(cmd);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            comboBox1.DataSource = dt;
            comboBox1.DisplayMember = "Số_CT";
            comboBox1.ValueMember = "Số_CT";
           //comboBox1.Text = dt.Rows[0][0].ToString();
            ////comboBox1.Text = comboBox1.SelectedText;
        }
        private void tinh_comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //load_tinh();
            //search_box_tinh();
            int fid;
            bool parseOK = Int32.TryParse(tinh_comboBox2.SelectedValue.ToString(), out fid);
            if (parseOK == true)
            {
                LoadHuyen(fid);
            }      
        }
        private void tinh_comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            //load_tinh();
            int fid;
            bool parseOK = Int32.TryParse(tinh_comboBox5.SelectedValue.ToString(), out fid);
            LoadHuyen_tt(fid);
        }
        private void huyen_comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            //int fid;
            //bool parseOK = Int32.TryParse(huyen_comboBox6.SelectedValue.ToString(), out fid);
            //LoadXa(fid);
        }
        private void huyen_comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            int fid;
            bool parseOK = Int32.TryParse(huyen_comboBox3.SelectedValue.ToString(), out fid);
            LoadXa(fid);
        }
        private void IN_BAO_ALL(string date)
        {
            if (cn.State != ConnectionState.Open)
            {
                cn.Open();
            }
            SqlCommand cmd = new SqlCommand("bao_cao_cthcc", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add(new SqlParameter("@ngayduyet", date));
            da = new SqlDataAdapter(cmd);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
            //Start Word and create a new document.
            Microsoft.Office.Interop.Word._Application oWord;
            Microsoft.Office.Interop.Word._Document oDoc;
            oWord = new Microsoft.Office.Interop.Word.Application();
            //oWord.PrintPreview = true;
            oWord.Width = 300;
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);
            ;
            //Insert a paragraph at the beginning of the document.
            Microsoft.Office.Interop.Word.Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Text = "Báo Cáo Nhanh Chuyển Trả Kết Quả Hành Chính Công Đã Duyệt";
            oPara1.Range.Font.Bold = 1;
            oPara1.Range.Font.Size = 20;
            oPara1.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter();
            //Insert a paragraph at the end of the document.
            Microsoft.Office.Interop.Word.Paragraph oPara2;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara2 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara2.Range.Text = "Ngày: " + dateTimePicker4.Value.ToShortDateString();
            oPara2.Format.SpaceAfter = 6;
            oPara2.Range.InsertParagraphAfter();
            //Insert another paragraph.
            Microsoft.Office.Interop.Word.Paragraph oPara3;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara3 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara3.Range.Font.Bold = 0;
            oPara3.Format.SpaceAfter = 24;
            oPara3.Range.InsertParagraphAfter();
            Microsoft.Office.Interop.Word.Table oTable;
            Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            int i = dt.Rows.Count;
            int j = dt.Columns.Count;
            oTable = oDoc.Tables.Add(wrdRng, i + 1, j, ref oMissing, ref oMissing); 
            oTable.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
            oTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            oTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            oTable.Cell(1, 1).Range.Text = "Tên Tỉnh/TP TƯ";
            oTable.Cell(1, 2).Range.Text = "Tên Huyện/Thành Phố";
            oTable.Cell(1, 3).Range.Text = "Số lượng";
            for (int k = 2; k <= oTable.Rows.Count; k++)
            {
                for (int l = 1; l <= oTable.Columns.Count; l++)
                {
                    oTable.Cell(k, l).Range.Text = dt.Rows[k - 2][l - 1].ToString();
                }
            }
            oTable.Rows[1].Range.Font.Bold = 1;
            oTable.Rows[1].Range.Font.Italic = 1;
            Microsoft.Office.Interop.Word.Paragraph oPara4;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara4 = oDoc.Content.Paragraphs.Add(ref oRng);
            oTable.Columns[1].Width = oWord.InchesToPoints(2);
            oTable.Columns[2].Width = oWord.InchesToPoints(2);
            oTable.Columns[3].Width = oWord.InchesToPoints(2);
            //Add text after the chart.
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            wrdRng.InsertParagraphAfter();
            wrdRng.InsertAfter("------------------------------");
            oDoc.Activate();
            oWord.Visible = true;
            oWord.PrintPreview = true;
            oWord.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintPreview;
            //object nullobj = Missing.Value;
            //int dialogResult = oWord.Dialogs[Microsoft.Office.Interop.Word.WdWordDialog.wdDialogFilePrint].Show(ref nullobj);
            //if (dialogResult == 1)
            //{
            //    oWord.PrintOut();
            //}
            cn.Close();
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //auto_add_id();
        }
        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkLabel2.LinkVisited = true;
            //Call the Process.Start method to open the default browser
            //with a URL:
            System.Diagnostics.Process.Start("https://www.facebook.com/tttiendanggg/");
        }
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkLabel1.LinkVisited = true;
            //Call the Process.Start method to open the default browser
            //with a URL:
            System.Diagnostics.Process.Start("https://github.com/tiendang218");
        }
        private void label7_Click(object sender, EventArgs e)
        {
        }
        void duyet_ct(int i)
        {
            if (DialogResult.Yes == MessageBox.Show("Xác nhận", "Cảnh báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                if (cn.State == ConnectionState.Closed)
                {
                    cn.Open();
                }
                try
                {
                    if (comboBox1.Text != string.Empty)
                    {
                        if (i == 1)
                        {
                            cmd = new SqlCommand("duyet_ct", cn);
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@id", int.Parse(comboBox1.Text));
                            cmd.Parameters.AddWithValue("@i", i);
                            cmd.Parameters.AddWithValue("@date", DateTime.Now.ToShortDateString());
                            cmd.ExecuteNonQuery();
                            duyet_box.Checked = true;
                            MessageBox.Show("Duyệt thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        if (i == 0)
                        {
                            cmd = new SqlCommand("duyet_ct", cn);
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@id", int.Parse(comboBox1.Text));
                            cmd.Parameters.AddWithValue("@i", i);
                            cmd.Parameters.AddWithValue("@date", "");
                            cmd.ExecuteNonQuery();
                            duyet_box.Checked = false;
                            MessageBox.Show("Hủy duyệt thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        Get_All_ChungTu();
                    }
                    else
                    {
                        MessageBox.Show("Điền thiếu dữ liệu hoặc số chứng từ đã tồn tại", "Invalid Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch
                { MessageBox.Show("Số CT bị trùng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                cn.Close();
            }
        }
        private void duyet_box_CheckedChanged(object sender, EventArgs e)
        {
        }
        private void duyet_box_CheckStateChanged(object sender, EventArgs e)
        {
        }
        private void in_banke_Click(object sender, EventArgs e)
        {
            switch(kieu_bao_cao)
            {
                case 1:
                    day = dateTimePicker4.Value.ToShortDateString();
                    Form fr2 = new Report();
                    fr2.Show();
                    break;
                case 2:
                    Form fr3 = new Report();
                    fr3.Show();
                    break;
                case 3:
                    day = dateTimePicker4.Value.ToShortDateString();
                    Form fr4 = new Report();
                    fr4.Show();
                    break;
                case 4:
                    day = dateTimePicker4.Value.ToShortDateString();
                    Form fr5 = new Report();
                    fr5.Show();
                    break;
                case 5:
                    day = dateTimePicker4.Value.ToShortDateString();
                    Form fr6 = new Report();
                    fr6.Show();
                    break;
                default:
                    break;
            }
        }
        private void printPreviewControl1_Click(object sender, EventArgs e)
        {
        }
        private void button1_Click(object sender, EventArgs e)
        {
            duyet_ct(1);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            duyet_ct(0);
        }
        private void so_hcc_box_TextChanged(object sender, EventArgs e)
        {
        }
        private void ten_box_TextChanged(object sender, EventArgs e)
        {
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                try
                {
                    if (cn.State != ConnectionState.Open)
                    {
                        cn.Open();
                    }
                    SqlCommand cmd = new SqlCommand("search_sdt", cn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@sdt", textBox1.Text));
                    da = new SqlDataAdapter(cmd);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
                }
                catch
                {
                    MessageBox.Show("Số chứng từ ko hợp lệ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                cn.Close();
            }
            else
            {
                Get_All_ChungTu();
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Get_All_ChungTu();
        }
        private void InBiThuDS_Click(object sender, EventArgs e)
        {
            //if (cn.State != ConnectionState.Open)
            //{
            //    cn.Open();
            //}
            //SqlCommand cmd = new SqlCommand("bao_cao_cthcc", cn);
            //cmd.CommandType = CommandType.StoredProcedure;
            //cmd.Parameters.Add(new SqlParameter("@ngayduyet", date));
            //da = new SqlDataAdapter(cmd);
            //System.Data.DataTable dt = new System.Data.DataTable();
            //da.Fill(dt);
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
            //Start Word and create a new document.
            Microsoft.Office.Interop.Word._Application oWord;
            Microsoft.Office.Interop.Word._Document oDoc;
            oWord = new Microsoft.Office.Interop.Word.Application();
            oDoc = new Microsoft.Office.Interop.Word.Document();
            
            //oWord.PrintPreview = true;
            oWord.Width = 300;
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);
            ;
            //Insert a paragraph at the beginning of the document.
            Microsoft.Office.Interop.Word.Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Text = "BƯU ĐIỆN TP.PHAN THIẾT";
            oPara1.Range.Text = "BƯU ĐIỆN TỈNH BÌNH THUẬN";
            oPara1.Range.Font.Bold = 1;
            oPara1.Range.Font.Size = 20;
            oPara1.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter();
            Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            wrdRng.InsertParagraphAfter();
            wrdRng.InsertAfter("Tên: "+ten_box.Text + "  SĐT: " + sdt_textBox2.Text +  "Địa chỉ thường trú: " + diachi_thuongtru_box.Text);
            wrdRng.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
            oDoc.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
            if (A5.CheckState == CheckState.Checked)
            {
                oDoc.PageSetup.PaperSize = WdPaperSize.wdPaperA5;
            }
            if (A6.CheckState == CheckState.Checked)
            {
                oDoc.PageSetup.PaperSize = WdPaperSize.wdPaperA3;
            }
            oDoc.Activate();
            oWord.Visible = true;
            oWord.PrintPreview = true;
            oWord.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintPreview;
            //object nullobj = Missing.Value;
            //int dialogResult = oWord.Dialogs[Microsoft.Office.Interop.Word.WdWordDialog.wdDialogFilePrint].Show(ref nullobj);
            //if (dialogResult == 1)
            //{
            //    oWord.PrintOut();
            //}
            cn.Close();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            //string _path;
            //OpenFileDialog od = new OpenFileDialog();
            //od.Filter = "Excell|*.xls;*.xlsx;";
            //od.FileName = "FileImport.xlsx";
            //BackgroundWorker bw = new BackgroundWorker
            //{
            //    WorkerReportsProgress = true,
            //    WorkerSupportsCancellation = true
            //};
            //DialogResult dr = od.ShowDialog();
            //if (dr == DialogResult.Abort)
            //    return;
            //if (dr == DialogResult.Cancel)
            //    return;
            ////txtpath.Text = od.FileName.ToString();   
            //if (dr == DialogResult.OK)
            //{
            //    try
            //    {
            //        _path = od.FileName.ToString();
            //        string path = _path;
            //        button4.Text = "Loading";
            //        button4.Enabled = false;
            //        if (bw.IsBusy)
            //        {
            //            return;
            //        }
            //        System.Diagnostics.Stopwatch sWatch = new System.Diagnostics.Stopwatch();
            //        bw.DoWork += (bwSender, bwArg) =>
            //        {
            //            //what happens here must not touch the form
            //            //as it's in a different thread
            //            sWatch.Start();
            //            System.Data.DataTable table = Exceldatatable(path);
            //            if (cn.State == ConnectionState.Closed)
            //            {
            //                cn.Open();
            //            }
            //            try
            //            {
            //                SqlCommand cmd = new SqlCommand("ImportTableExcel", cn);
            //                cmd.CommandType = CommandType.StoredProcedure;
            //                SqlParameter dtparam = cmd.Parameters.AddWithValue("@table_excel", table);
            //                dtparam.SqlDbType = SqlDbType.Structured;
            //                cmd.ExecuteNonQuery();
            //            }
            //            catch
            //            {
            //                MessageBox.Show("Lỗi dữ liệu", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //            }
            //        };
            //        bw.ProgressChanged += (bwSender, bwArg) =>
            //            {
            //            };
            //        bw.RunWorkerCompleted += (bwSender, bwArg) =>
            //            {
            //                sWatch.Stop();
            //                Get_All_ChungTu();
            //                button4.Enabled = true;
            //                button4.Text = "Nhập Excel";
            //                bw.Dispose();
            //                cn.Close();
            //            };
            //        //Starts the actual work - triggerrs the "DoWork" event
            //        bw.RunWorkerAsync();
            //    }
            //    catch
            //    {
            //        MessageBox.Show("Dữ liệu file excel quá lớn hoặc quá thời gian truy vấn", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    }
            //}
        }
        public static System.Data.DataTable Exceldatatable(string path)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            try
            {
                string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;MAXSCANROWS=0'";
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    using (OleDbCommand comm = new OleDbCommand())
                    {
                        string sheetName = "Sheet1";
                        comm.CommandText = "Select * from [" + sheetName + "$]";
                        comm.Connection = conn;
                        using (OleDbDataAdapter da = new OleDbDataAdapter())
                        {
                            da.SelectCommand = comm;
                            da.Fill(dt);
                            return dt;
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("File Excel đang mở, tắt để đọc file", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (textBox2.Text != "")
            {
                try
                {
                    if (cn.State != ConnectionState.Open)
                    {
                        cn.Open();
                    }
                    SqlCommand cmd = new SqlCommand("search_nguoinhan", cn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@ten", textBox2.Text));
                    da = new SqlDataAdapter(cmd);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
                }
                catch
                {
                    MessageBox.Show("Không tìm thấy tên người nhận", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                cn.Close();
            }
            else
            {
                Get_All_ChungTu();
            }
        }
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.Value = dateTimePicker1.Value.AddDays(7);
        }
        private void Form1_Click(object sender, EventArgs e)
        {
            Get_All_ChungTu();
            Them.Enabled = true;
            auto_add_id();
            xoa_noidung_box();
        }
        private void Cach_Bao_Cao_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selected = Cach_Bao_Cao.GetItemText(Cach_Bao_Cao.SelectedItem);
            switch(selected)
            {
                case "Theo ngày nhận":
                    kieu_bao_cao = 1;
                    dateTimePicker4.Visible = true;
                    label24.Visible = true;
                    break;
                case "Tất cả ngày":
                    kieu_bao_cao = 2;
                    dateTimePicker4.Visible = false;
                    label24.Visible = false;
                    break;
                case "Theo ngày duyệt":
                    kieu_bao_cao = 3;
                    dateTimePicker4.Visible = true;
                    label24.Visible = true;
                    break;
                case "Theo ngày hẹn" :
                    kieu_bao_cao = 4;
                    dateTimePicker4.Visible = true;
                    label24.Visible = true;
                    break;
                case "Theo ngày nhập":
                    kieu_bao_cao = 5;
                    dateTimePicker4.Visible = true;
                    label24.Visible = true;
                    break;
                default:
                    kieu_bao_cao = 0;
                    dateTimePicker4.Visible = false;
                    label24.Visible = false;
                    break;
            }
        }
        private void dateTimePicker4_ValueChanged(object sender, EventArgs e)
        {
            Form1.day = dateTimePicker4.Value.ToShortDateString();
        }
        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void comboBox1_Click(object sender, EventArgs e)
        {
            //auto_add_id();
        }

        private void A5_CheckedChanged(object sender, EventArgs e)
        {
            if (A5.Checked == true)
            {
                A6.CheckState = CheckState.Unchecked;
            }
            else
                A6.CheckState = CheckState.Checked;
        }

        private void A6_CheckedChanged(object sender, EventArgs e)
        {
            if (A6.Checked == true)
            {
                A5.CheckState = CheckState.Unchecked;
            }
            else
                A5.CheckState = CheckState.Checked;
        }

        private void label18_Click(object sender, EventArgs e)
        {

        }
    }
}