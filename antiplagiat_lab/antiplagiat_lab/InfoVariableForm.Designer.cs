namespace antiplagiat_lab
{
  partial class InfoVariableForm
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
      this.rTbxInfoVariable = new System.Windows.Forms.RichTextBox();
      this.splitter1 = new System.Windows.Forms.Splitter();
      this.SuspendLayout();
      // 
      // rTbxInfoVariable
      // 
      this.rTbxInfoVariable.Location = new System.Drawing.Point(0, 66);
      this.rTbxInfoVariable.Name = "rTbxInfoVariable";
      this.rTbxInfoVariable.Size = new System.Drawing.Size(800, 384);
      this.rTbxInfoVariable.TabIndex = 0;
      this.rTbxInfoVariable.Text = "";
      // 
      // splitter1
      // 
      this.splitter1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.splitter1.Dock = System.Windows.Forms.DockStyle.Top;
      this.splitter1.Location = new System.Drawing.Point(0, 0);
      this.splitter1.Name = "splitter1";
      this.splitter1.Size = new System.Drawing.Size(800, 10);
      this.splitter1.TabIndex = 2;
      this.splitter1.TabStop = false;
      // 
      // InfoVariableForm
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(800, 450);
      this.Controls.Add(this.splitter1);
      this.Controls.Add(this.rTbxInfoVariable);
      this.Name = "InfoVariableForm";
      this.Text = "Сравнение переменных";
      this.Load += new System.EventHandler(this.InfoVariableForm_Load);
      this.ResumeLayout(false);

    }

    #endregion

    private System.Windows.Forms.RichTextBox rTbxInfoVariable;
        private System.Windows.Forms.Splitter splitter1;
    }
}