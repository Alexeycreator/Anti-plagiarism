namespace antiplagiat_lab
{
    partial class ReportForm
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
            this.listBox_Reports = new System.Windows.Forms.ListBox();
            this.button_DeleteReport = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // listBox_Reports
            // 
            this.listBox_Reports.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.listBox_Reports.FormattingEnabled = true;
            this.listBox_Reports.ItemHeight = 24;
            this.listBox_Reports.Location = new System.Drawing.Point(16, 46);
            this.listBox_Reports.Name = "listBox_Reports";
            this.listBox_Reports.Size = new System.Drawing.Size(238, 172);
            this.listBox_Reports.TabIndex = 0;
            // 
            // button_DeleteReport
            // 
            this.button_DeleteReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.button_DeleteReport.Location = new System.Drawing.Point(16, 224);
            this.button_DeleteReport.Name = "button_DeleteReport";
            this.button_DeleteReport.Size = new System.Drawing.Size(127, 39);
            this.button_DeleteReport.TabIndex = 1;
            this.button_DeleteReport.Text = "Удалить";
            this.button_DeleteReport.UseVisualStyleBackColor = true;
            this.button_DeleteReport.Click += new System.EventHandler(this.buttonDeleteReport_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.label3.Location = new System.Drawing.Point(12, 19);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(159, 24);
            this.label3.TabIndex = 8;
            this.label3.Text = "Список отчётов:";
            // 
            // ReportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(278, 277);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.button_DeleteReport);
            this.Controls.Add(this.listBox_Reports);
            this.MinimumSize = new System.Drawing.Size(294, 316);
            this.Name = "ReportForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Список Отчётов";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox listBox_Reports;
        private System.Windows.Forms.Button button_DeleteReport;
        private System.Windows.Forms.Label label3;
    }
}