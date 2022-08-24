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
            this.report_by_huyenTableAdapter1 = new XuatExcelApp.DataSet3TableAdapters.report_by_huyenTableAdapter();
            this.report_by_huyenTableAdapter2 = new XuatExcelApp.DataSet3TableAdapters.report_by_huyenTableAdapter();
            this.report_by_huyenTableAdapter3 = new XuatExcelApp.DataSet3TableAdapters.report_by_huyenTableAdapter();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.dataSet3 = new XuatExcelApp.DataSet3();
            this.reportbyhuyenBindingSource = new System.Windows.Forms.BindingSource(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.dataSet3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.reportbyhuyenBindingSource)).BeginInit();
            this.SuspendLayout();
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
            this.reportViewer1.Size = new System.Drawing.Size(1669, 814);
            this.reportViewer1.TabIndex = 0;
            this.reportViewer1.Load += new System.EventHandler(this.reportViewer1_Load);
            // 
            // dataSet3
            // 
            this.dataSet3.DataSetName = "DataSet3";
            this.dataSet3.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // reportbyhuyenBindingSource
            // 
            this.reportbyhuyenBindingSource.DataMember = "report_by_huyen";
            this.reportbyhuyenBindingSource.DataSource = this.dataSet3;
            // 
            // Report
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(1693, 838);
            this.Controls.Add(this.reportViewer1);
            this.Name = "Report";
            this.Text = "HCC Report";
            this.Load += new System.EventHandler(this.Report_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataSet3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.reportbyhuyenBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DataSet3TableAdapters.report_by_huyenTableAdapter report_by_huyenTableAdapter1;
        private DataSet3TableAdapters.report_by_huyenTableAdapter report_by_huyenTableAdapter2;
        private DataSet3TableAdapters.report_by_huyenTableAdapter report_by_huyenTableAdapter3;
        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.BindingSource reportbyhuyenBindingSource;
        private DataSet3 dataSet3;
    }
}