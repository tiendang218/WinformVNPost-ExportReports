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
            this.report_by_huyenTableAdapter1 = new XuatExcelApp.ChungTuHCCDataSetTableAdapters.report_by_huyenTableAdapter();
            this.report_by_huyenTableAdapter2 = new XuatExcelApp.ChungTuHCCDataSetTableAdapters.report_by_huyenTableAdapter();
            this.report_by_huyenTableAdapter3 = new XuatExcelApp.ChungTuHCCDataSetTableAdapters.report_by_huyenTableAdapter();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
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
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "XuatExcelApp.Report1.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(12, 12);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.ServerReport.BearerToken = null;
            this.reportViewer1.Size = new System.Drawing.Size(1412, 565);
            this.reportViewer1.TabIndex = 0;
            this.reportViewer1.Load += new System.EventHandler(this.reportViewer1_Load);
            // 
            // Report
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(1436, 589);
            this.Controls.Add(this.reportViewer1);
            this.Name = "Report";
            this.Text = "HCC Report";
            this.Load += new System.EventHandler(this.Report_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ChungTuHCCDataSetTableAdapters.report_by_huyenTableAdapter report_by_huyenTableAdapter1;
        private ChungTuHCCDataSetTableAdapters.report_by_huyenTableAdapter report_by_huyenTableAdapter2;
        private ChungTuHCCDataSetTableAdapters.report_by_huyenTableAdapter report_by_huyenTableAdapter3;
        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
    }
}