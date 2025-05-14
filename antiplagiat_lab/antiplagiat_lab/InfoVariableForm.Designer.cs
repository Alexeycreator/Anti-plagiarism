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
      this.SuspendLayout();
      // 
      // rTbxInfoVariable
      // 
      this.rTbxInfoVariable.Location = new System.Drawing.Point(12, 12);
      this.rTbxInfoVariable.Name = "rTbxInfoVariable";
      this.rTbxInfoVariable.Size = new System.Drawing.Size(782, 426);
      this.rTbxInfoVariable.TabIndex = 0;
      this.rTbxInfoVariable.Text = "";
      // 
      // InfoVariableForm
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(800, 450);
      this.Controls.Add(this.rTbxInfoVariable);
      this.Name = "InfoVariableForm";
      this.Text = "InfoVariableForm";
      this.Load += new System.EventHandler(this.InfoVariableForm_Load);
      this.ResumeLayout(false);

    }

    #endregion

    private System.Windows.Forms.RichTextBox rTbxInfoVariable;
  }
}