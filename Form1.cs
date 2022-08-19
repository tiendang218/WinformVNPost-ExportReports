using System;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Spire.Doc;
namespace XuatExcelApp
{

    public partial class Form1 : Form
    {
        SqlConnection cn = new SqlConnection(@"Data Source=KTNV-TIEN\SQLEXPRESS;Initial Catalog=ChungTuHCC;Integrated Security=True");
        SqlCommand cmd;
        SqlDataAdapter da;
        SqlDataReader dr;
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
            load_loai_hoso();
            LoadHuyen(int.Parse(tinh_comboBox2.SelectedValue.ToString()));
            LoadXa(int.Parse(huyen_comboBox3.SelectedValue.ToString()));
            ////disable delete and update button on load  
            auto_add_id();
            Sửa.Enabled = true;
            Xóa.Enabled = true;
            cn.Close();
        }
        private void Get_All_ChungTu()
        {
            cmd = new SqlCommand("CRUD", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@SO_CT", 0);
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
                        cmd.Parameters.AddWithValue("@TEN_NGUOI_GUI", ten_box.Text);
                        cmd.Parameters.AddWithValue("@DIA_CHI", diachi_box.Text);
                        cmd.Parameters.AddWithValue("@SO_HS_HCC", so_hcc_box.Text);
                        cmd.Parameters.AddWithValue("@NGAY_HEN", dateTimePicker2.Value);
                        cmd.Parameters.AddWithValue("@NGAY_NHAN", dateTimePicker1.Value);
                        cmd.Parameters.AddWithValue("@TINH_NHAN", int.Parse(tinh_comboBox2.SelectedValue.ToString()));
                        cmd.Parameters.AddWithValue("@HUYEN_NHAN", int.Parse(huyen_comboBox3.SelectedValue.ToString()));
                        cmd.Parameters.AddWithValue("@XA_NHAN", int.Parse(xa_comboBox4.SelectedValue.ToString()));
                        cmd.Parameters.AddWithValue("@CUOC", cuoc_textBox8.Text);
                        cmd.Parameters.AddWithValue("@SO_HS_KEM", sohskem.Text);
                        cmd.Parameters.AddWithValue("@DIEN_THOAI", sdt_textBox2.Text);
                        cmd.Parameters.AddWithValue("@MA_LOAI_HS", int.Parse(loai_hs_comboBox.SelectedValue.ToString()));
                        cmd.Parameters.AddWithValue("@TRONG_LUONG", trongluong.Text);
                        cmd.Parameters.AddWithValue("@MA_BUUGUI", mabuugui.Text);
                        cmd.Parameters.AddWithValue("@GHI_CHU", ghichu_textBox9.Text);
                        cmd.Parameters.AddWithValue("@OperationType", "1");
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Bản ghi được thêm thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Get_All_ChungTu();
                    }
                    else
                    {
                        MessageBox.Show("Điền thiếu dữ liệu", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }

                catch
                { MessageBox.Show("Số CT bị trùng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information); }

            
        
}
        
        private void load_tinh()
        {
            cmd = new SqlCommand("CRUD", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@SO_CT", 0);
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
            cmd.Parameters.AddWithValue("@SO_CT", 0);
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
        void LoadHuyen(int MaTinh)
        {
            string sql = @"select * from HUYEN_TP where MA_TINH = @MaTinh";   
            SqlCommand cmd = new SqlCommand(sql, cn);
            SqlParameter para = new SqlParameter("@MaTinh", SqlDbType.Int);
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
        void LoadXa(int MaHuyen)
        {
            string sql = @"select * from XAPHUONG where MA_HUYENTP= @MaHuyen";
            SqlCommand cmd = new SqlCommand(sql, cn);
            SqlParameter para = new SqlParameter("@MaHuyen", SqlDbType.Int);
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
            int x = int.Parse(tinh_comboBox5.SelectedIndex.ToString());
            LoadHuyen(86);

        }

      

        private void tinh_comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string x = (tinh_comboBox2.SelectedValue).ToString();
            LoadHuyen(86);
        }
        private void huyen_comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            int x = int.Parse(huyen_comboBox3.SelectedIndex.ToString());
            LoadXa(1);

        }

        private void huyen_comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            int x = int.Parse(huyen_comboBox6.SelectedIndex.ToString());
            LoadXa(1);

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void Xóa_Click(object sender, EventArgs e)
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
                cmd.Parameters.AddWithValue("@TEN_NGUOI_GUI", ten_box.Text);
                cmd.Parameters.AddWithValue("@DIA_CHI", diachi_box.Text);
                cmd.Parameters.AddWithValue("@SO_HS_HCC", so_hcc_box.Text);
                cmd.Parameters.AddWithValue("@NGAY_HEN", dateTimePicker2.Value);
                cmd.Parameters.AddWithValue("@NGAY_NHAN", dateTimePicker1.Value);
                cmd.Parameters.AddWithValue("@TINH_NHAN", int.Parse(tinh_comboBox2.SelectedValue.ToString()));
                cmd.Parameters.AddWithValue("@HUYEN_NHAN", int.Parse(huyen_comboBox3.SelectedValue.ToString()));
                cmd.Parameters.AddWithValue("@XA_NHAN", int.Parse(xa_comboBox4.SelectedValue.ToString()));
                cmd.Parameters.AddWithValue("@CUOC", cuoc_textBox8.Text);
                cmd.Parameters.AddWithValue("@SO_HS_KEM", sohskem.Text);
                cmd.Parameters.AddWithValue("@DIEN_THOAI", sdt_textBox2.Text);
                cmd.Parameters.AddWithValue("@MA_LOAI_HS", int.Parse(loai_hs_comboBox.SelectedValue.ToString()));
                cmd.Parameters.AddWithValue("@TRONG_LUONG", trongluong.Text);
                cmd.Parameters.AddWithValue("@MA_BUUGUI", mabuugui.Text);
                cmd.Parameters.AddWithValue("@GHI_CHU", ghichu_textBox9.Text);
                cmd.Parameters.AddWithValue("@OperationType", "3");
                cmd.ExecuteNonQuery();
                MessageBox.Show("Bản ghi được xóa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Get_All_ChungTu();
            }
            catch
            { MessageBox.Show("Số CT bị trùng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information); }


        }
       private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {   
            
            comboBox1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
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
            comboBox1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            //so_ct_textbox.Enabled = false;    
        }
        private void Sửa_Click(object sender, EventArgs e)
        {
            if (cn.State == ConnectionState.Closed)
            {
                cn.Open();
            }
            try
            {

                if (comboBox1.Text != string.Empty)
                {
                    cmd = new SqlCommand("CRUD", cn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@SO_CT", int.Parse(comboBox1.Text));
                    cmd.Parameters.AddWithValue("@TEN_NGUOI_GUI", ten_box.Text);
                    cmd.Parameters.AddWithValue("@DIA_CHI", diachi_box.Text);
                    cmd.Parameters.AddWithValue("@SO_HS_HCC", so_hcc_box.Text);
                    cmd.Parameters.AddWithValue("@NGAY_HEN", dateTimePicker2.Value);
                    cmd.Parameters.AddWithValue("@NGAY_NHAN", dateTimePicker1.Value);
                    cmd.Parameters.AddWithValue("@TINH_NHAN", int.Parse(tinh_comboBox2.SelectedValue.ToString()));
                    cmd.Parameters.AddWithValue("@HUYEN_NHAN", int.Parse(huyen_comboBox3.SelectedValue.ToString()));
                    cmd.Parameters.AddWithValue("@XA_NHAN", int.Parse(xa_comboBox4.SelectedValue.ToString()));
                    cmd.Parameters.AddWithValue("@CUOC", cuoc_textBox8.Text);
                    cmd.Parameters.AddWithValue("@SO_HS_KEM", sohskem.Text);
                    cmd.Parameters.AddWithValue("@DIEN_THOAI", sdt_textBox2.Text);
                    cmd.Parameters.AddWithValue("@MA_LOAI_HS", int.Parse(loai_hs_comboBox.SelectedValue.ToString()));
                    cmd.Parameters.AddWithValue("@TRONG_LUONG", trongluong.Text);
                    cmd.Parameters.AddWithValue("@MA_BUUGUI", mabuugui.Text);
                    cmd.Parameters.AddWithValue("@GHI_CHU", ghichu_textBox9.Text);
                    cmd.Parameters.AddWithValue("@OperationType", "2");
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Bản ghi được sửa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Get_All_ChungTu();
                    //so_ct_textbox.Text = "";
                    //ten_box.Text = "";
                    //diachi_box.Text = "";
                    //so_hcc_box.Text = "";
                    //dateTimePicker2.Text = "";
                    //dateTimePicker1.Text = "";
                    //tinh_comboBox2.Text = "";
                    //huyen_comboBox3.Text = "";
                    //xa_comboBox4.Text = "";
                    //cuoc_textBox8.Text = "";
                    //sohskem.Text = "";
                    //sdt_textBox2.Text = "";
                    //loai_hs_comboBox.Text = "";
                    //trongluong.Text = "";
                    //mabuugui.Text = "";
                    //ghichu_textBox9.Text = "";
                    //Xóa.Enabled = false;
                }
                else
                {
                    MessageBox.Show("Điền thiếu dữ liệu hoặc số chứng từ đã tồn tại", "Invalid Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch
            { MessageBox.Show("Số CT bị trùng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information); }
            
        }
        private void xuatExcel_Click(object sender, EventArgs e)
        {
            cn.Open();
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
            worksheet.Columns.ColumnWidth = 15;
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
            cn.Close();
        }
        private void button7_Click(object sender, EventArgs e)
        {
            
            string ten_huyen = huyen_comboBox3.Text;
            tenhuyen_report.Text = ten_huyen;
            IN_BAOCAO(int.Parse((tinh_comboBox5.SelectedValue).ToString()), int.Parse((huyen_comboBox6.SelectedValue).ToString()));
            
        }
        void IN_BAOCAO(int ma_tinh, int ma_huyen)
        {
            cn.Open();
            SqlCommand cmd = new SqlCommand("dem_ban_ghi", cn); 
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add(new SqlParameter("@tinh", ma_tinh));
            cmd.Parameters.Add(new SqlParameter("@huyen", ma_huyen));
            //Add parameters like this
            int count = (int)cmd.ExecuteScalar();
            baocao.ResetText();
            baocao.Text = count.ToString();

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
            oPara1.Range.Text = "Báo Cáo Nhanh Chuyển Trả Kết Quả Hành Chính Công Tại " + huyen_comboBox3.Text.ToString(); ;
            oPara1.Range.Font.Bold = 1;
            oPara1.Range.Font.Size = 20;
            oPara1.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter();

            //Insert a paragraph at the end of the document.
            Microsoft.Office.Interop.Word.Paragraph oPara2;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara2 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara2.Range.Text = "Ngày: " + DateTime.Today.ToString();
            oPara2.Format.SpaceAfter = 6;
            oPara2.Range.InsertParagraphAfter();

            //Insert another paragraph.
            Microsoft.Office.Interop.Word.Paragraph oPara3;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara3 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara3.Range.Font.Bold = 0;
            oPara3.Format.SpaceAfter = 24;
            oPara3.Range.InsertParagraphAfter();

            //Insert a 3 x 3 table, fill it with data, and make the first row
            //bold and italic.
            Microsoft.Office.Interop.Word.Table oTable;
            Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            oTable = oDoc.Tables.Add(wrdRng, 3, 3, ref oMissing, ref oMissing);
            oTable.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
            oTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            oTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            oTable.Cell(1, 1).Range.Text = "STT";
            oTable.Cell(1, 2).Range.Text = "Tên Đơn Vị";
            oTable.Cell(1, 3).Range.Text = "Số lượng";
            oTable.Cell(2, 1).Range.Text = "1";
            oTable.Cell(2, 2).Range.Text = huyen_comboBox3.Text.ToString();
            oTable.Cell(2, 3).Range.Text = count.ToString();
            oTable.Cell(3, 2).Range.Text = "Tổng cộng";
            oTable.Cell(3, 3).Range.Text = count.ToString();

            oTable.Rows[1].Range.Font.Bold = 1;
            oTable.Rows[1].Range.Font.Italic = 1;

            //Add some text after the table.
            Microsoft.Office.Interop.Word.Paragraph oPara4;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara4 = oDoc.Content.Paragraphs.Add(ref oRng);
            oTable.Columns[1].Width = oWord.InchesToPoints(1); //Change width of columns 1 & 2
            oTable.Columns[2].Width = oWord.InchesToPoints(4);
            oTable.Columns[3].Width = oWord.InchesToPoints(2);
            //Add text after the chart.
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            wrdRng.InsertParagraphAfter();
            wrdRng.InsertAfter("------------------------------");
            oWord.Visible = true;
            //object nullobj = Missing.Value;
            //int dialogResult = oWord.Dialogs[Microsoft.Office.Interop.Word.WdWordDialog.wdDialogFilePrint].Show(ref nullobj);

            //if (dialogResult == 1)
            //{
            //    oWord.PrintOut();
            //}
            cn.Close();
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

        private void huyen_comboBox6_SelectedIndexChanged_1(object sender, EventArgs e)
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
            comboBox1.DisplayMember = "Số CT";
            comboBox1.ValueMember = "Số_CT";
            comboBox1.Text = comboBox1.ValueMember.ToString();
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            
        }

        private void comboBox1_TextUpdate(object sender, EventArgs e)
        {
            auto_add_id();
        }
    }
}
