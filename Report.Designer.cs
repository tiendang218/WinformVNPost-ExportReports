namespace XuatExcelApp
{
    partial class Report
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource1 = new Microsoft.Reporting.WinForms.ReportDataSource();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Report));
            this.reportbyhuyenBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.dataSet1 = new XuatExcelApp.DataSet1();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.report_by_huyenTableAdapter1 = new XuatExcelApp.DataSet1TableAdapters.report_by_huyenTableAdapter();
            this.report_by_huyenTableAdapter2 = new XuatExcelApp.DataSet1TableAdapters.report_by_huyenTableAdapter();
            this.report_by_huyenTableAdapter3 = new XuatExcelApp.DataSet1TableAdapters.report_by_huyenTableAdapter();
            this.report_by_huyenBindingSource = new System.Windows.Forms.BindingSource(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.reportbyhuyenBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.report_by_huyenBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // reportbyhuyenBindingSource
            // 
            this.reportbyhuyenBindingSource.DataMember = "report_by_huyen";
            this.reportbyhuyenBindingSource.DataSource = this.dataSet1;
            // 
            // dataSet1
            // 
            this.dataSet1.DataSetName = "DataSet1";
            this.dataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // reportViewer1
            // 
            this.reportViewer1.AutoScroll = true;
            this.reportViewer1.AutoSize = true;
            reportDataSource1.Name = "DataSet1";
            reportDataSource1.Value = this.reportbyhuyenBindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource1);
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "XuatExcelApp.Report1.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(12, 12);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.ServerReport.BearerToken = null;
            this.reportViewer1.Size = new System.Drawing.Size(1858, 873);
            this.reportViewer1.TabIndex = 0;
            this.reportViewer1.Load += new System.EventHandler(this.reportViewer1_Load);
            // 
            // report_by_huyenTableAdapter1
            // 
            this.report_by_huyenTableAdapter1.ClearBeforeFill = true;
            // 
            // report_by_huyenTableAdapter2
            // 
            this.report_by_huyenTableAdapter2.ClearBeforeFill = true;
            // 
            // report_by_huyenTableAdapter3
            // 
            this.report_by_huyenTableAdapter3.ClearBeforeFill = true;
            // 
            // report_by_huyenBindingSource
            // 
            this.report_by_huyenBindingSource.DataMember = "report_by_huyen";
            this.report_by_huyenBindingSource.DataSource = this.dataSet1;
            // 
            // Report
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(1882, 897);
            this.Controls.Add(this.reportViewer1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Report";
            this.Text = "Báo Cáo Chứng Từ Hành Chính Công";
            this.Load += new System.EventHandler(this.Report_Load);
            ((System.ComponentModel.ISupportInitialize)(this.reportbyhuyenBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.report_by_huyenBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DataSet1TableAdapters.report_by_huyenTableAdapter report_by_huyenTableAdapter1;
        private DataSet1TableAdapters.report_by_huyenTableAdapter report_by_huyenTableAdapter2;
        private DataSet1TableAdapters.report_by_huyenTableAdapter report_by_huyenTableAdapter3;
        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.BindingSource reportbyhuyenBindingSource;
        private DataSet1 dataSet1;
        private System.Windows.Forms.BindingSource report_by_huyenBindingSource;
    }
}