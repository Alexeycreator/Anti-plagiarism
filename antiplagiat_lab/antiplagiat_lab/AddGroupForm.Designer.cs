namespace antiplagiat_lab
{
    partial class AddGroupForm
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
            this.textBox_NewStudent = new System.Windows.Forms.TextBox();
            this.textBox_GroupName = new System.Windows.Forms.TextBox();
            this.listBox_Students = new System.Windows.Forms.ListBox();
            this.button_add = new System.Windows.Forms.Button();
            this.button_save = new System.Windows.Forms.Button();
            this.button_cancel = new System.Windows.Forms.Button();
            this.buttondelete = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btn_load_st = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox_NewStudent
            // 
            this.textBox_NewStudent.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.textBox_NewStudent.Location = new System.Drawing.Point(17, 165);
            this.textBox_NewStudent.Name = "textBox_NewStudent";
            this.textBox_NewStudent.Size = new System.Drawing.Size(241, 29);
            this.textBox_NewStudent.TabIndex = 0;
            // 
            // textBox_GroupName
            // 
            this.textBox_GroupName.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.textBox_GroupName.Location = new System.Drawing.Point(16, 63);
            this.textBox_GroupName.Name = "textBox_GroupName";
            this.textBox_GroupName.Size = new System.Drawing.Size(161, 29);
            this.textBox_GroupName.TabIndex = 0;
            // 
            // listBox_Students
            // 
            this.listBox_Students.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listBox_Students.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.listBox_Students.FormattingEnabled = true;
            this.listBox_Students.ItemHeight = 24;
            this.listBox_Students.Location = new System.Drawing.Point(356, 36);
            this.listBox_Students.Name = "listBox_Students";
            this.listBox_Students.Size = new System.Drawing.Size(238, 172);
            this.listBox_Students.TabIndex = 1;
            // 
            // button_add
            // 
            this.button_add.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.button_add.Location = new System.Drawing.Point(12, 214);
            this.button_add.Name = "button_add";
            this.button_add.Size = new System.Drawing.Size(127, 39);
            this.button_add.TabIndex = 2;
            this.button_add.Text = "Добавить";
            this.button_add.UseVisualStyleBackColor = true;
            this.button_add.Click += new System.EventHandler(this.buttonAddStudent_Click);
            // 
            // button_save
            // 
            this.button_save.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button_save.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.button_save.Location = new System.Drawing.Point(356, 299);
            this.button_save.Name = "button_save";
            this.button_save.Size = new System.Drawing.Size(127, 39);
            this.button_save.TabIndex = 2;
            this.button_save.Text = "Сохранить";
            this.button_save.UseVisualStyleBackColor = true;
            this.button_save.Click += new System.EventHandler(this.buttonSave_Click);
            // 
            // button_cancel
            // 
            this.button_cancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button_cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.button_cancel.Location = new System.Drawing.Point(489, 299);
            this.button_cancel.Name = "button_cancel";
            this.button_cancel.Size = new System.Drawing.Size(127, 39);
            this.button_cancel.TabIndex = 2;
            this.button_cancel.Text = "Отмена";
            this.button_cancel.UseVisualStyleBackColor = true;
            this.button_cancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // buttondelete
            // 
            this.buttondelete.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.buttondelete.Location = new System.Drawing.Point(356, 214);
            this.buttondelete.Name = "buttondelete";
            this.buttondelete.Size = new System.Drawing.Size(127, 39);
            this.buttondelete.TabIndex = 2;
            this.buttondelete.Text = "Удалить";
            this.buttondelete.UseVisualStyleBackColor = true;
            this.buttondelete.Click += new System.EventHandler(this.buttonDeleteStudent_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.label1.Location = new System.Drawing.Point(12, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(171, 24);
            this.label1.TabIndex = 3;
            this.label1.Text = "Название группы:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.label2.Location = new System.Drawing.Point(13, 138);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(147, 24);
            this.label2.TabIndex = 3;
            this.label2.Text = "ФИО студента:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.label3.Location = new System.Drawing.Point(352, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(149, 24);
            this.label3.TabIndex = 3;
            this.label3.Text = "Список группы:";
            // 
            // btn_load_st
            // 
            this.btn_load_st.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.btn_load_st.Location = new System.Drawing.Point(12, 280);
            this.btn_load_st.Name = "btn_load_st";
            this.btn_load_st.Size = new System.Drawing.Size(127, 39);
            this.btn_load_st.TabIndex = 4;
            this.btn_load_st.Text = "Загрузить из файла";
            this.btn_load_st.UseVisualStyleBackColor = true;
            this.btn_load_st.Click += new System.EventHandler(this.btn_load_st_Click);
            // 
            // AddGroupForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(628, 350);
            this.Controls.Add(this.btn_load_st);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button_cancel);
            this.Controls.Add(this.buttondelete);
            this.Controls.Add(this.button_save);
            this.Controls.Add(this.button_add);
            this.Controls.Add(this.listBox_Students);
            this.Controls.Add(this.textBox_GroupName);
            this.Controls.Add(this.textBox_NewStudent);
            this.MinimumSize = new System.Drawing.Size(644, 389);
            this.Name = "AddGroupForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Добавление новой группы";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox_NewStudent;
        private System.Windows.Forms.TextBox textBox_GroupName;
        private System.Windows.Forms.ListBox listBox_Students;
        private System.Windows.Forms.Button button_add;
        private System.Windows.Forms.Button button_save;
        private System.Windows.Forms.Button button_cancel;
        private System.Windows.Forms.Button buttondelete;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btn_load_st;
    }
}