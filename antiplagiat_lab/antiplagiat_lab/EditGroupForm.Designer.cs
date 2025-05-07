using System.Collections.Generic;

namespace antiplagiat_lab
{
    partial class EditGroupForm
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
            this.comboBox_Groups = new System.Windows.Forms.ComboBox();
            this.listBox_Students = new System.Windows.Forms.ListBox();
            this.textBox_NewStudent = new System.Windows.Forms.TextBox();
            this.button_add = new System.Windows.Forms.Button();
            this.button_remove = new System.Windows.Forms.Button();
            this.button_save = new System.Windows.Forms.Button();
            this.button_cancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btn_upload_file = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // comboBox_Groups
            // 
            this.comboBox_Groups.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.comboBox_Groups.FormattingEnabled = true;
            this.comboBox_Groups.Location = new System.Drawing.Point(16, 63);
            this.comboBox_Groups.Name = "comboBox_Groups";
            this.comboBox_Groups.Size = new System.Drawing.Size(161, 32);
            this.comboBox_Groups.TabIndex = 0;
            this.comboBox_Groups.SelectedIndexChanged += new System.EventHandler(this.comboBoxGroups_SelectedIndexChanged);
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
            // textBox_NewStudent
            // 
            this.textBox_NewStudent.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.textBox_NewStudent.Location = new System.Drawing.Point(17, 165);
            this.textBox_NewStudent.Name = "textBox_NewStudent";
            this.textBox_NewStudent.Size = new System.Drawing.Size(241, 29);
            this.textBox_NewStudent.TabIndex = 2;
            // 
            // button_add
            // 
            this.button_add.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.button_add.Location = new System.Drawing.Point(12, 214);
            this.button_add.Name = "button_add";
            this.button_add.Size = new System.Drawing.Size(127, 39);
            this.button_add.TabIndex = 3;
            this.button_add.Text = "Добавить";
            this.button_add.UseVisualStyleBackColor = true;
            this.button_add.Click += new System.EventHandler(this.buttonAddStudent_Click);
            // 
            // button_remove
            // 
            this.button_remove.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.button_remove.Location = new System.Drawing.Point(356, 214);
            this.button_remove.Name = "button_remove";
            this.button_remove.Size = new System.Drawing.Size(127, 39);
            this.button_remove.TabIndex = 3;
            this.button_remove.Text = "Удалить";
            this.button_remove.UseVisualStyleBackColor = true;
            this.button_remove.Click += new System.EventHandler(this.buttonRemoveStudent_Click);
            // 
            // button_save
            // 
            this.button_save.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button_save.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.button_save.Location = new System.Drawing.Point(356, 299);
            this.button_save.Name = "button_save";
            this.button_save.Size = new System.Drawing.Size(127, 39);
            this.button_save.TabIndex = 4;
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
            this.button_cancel.TabIndex = 4;
            this.button_cancel.Text = "Отменить";
            this.button_cancel.UseVisualStyleBackColor = true;
            this.button_cancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.label1.Location = new System.Drawing.Point(12, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(170, 24);
            this.label1.TabIndex = 5;
            this.label1.Text = "Выберите группу:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.label2.Location = new System.Drawing.Point(13, 138);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(147, 24);
            this.label2.TabIndex = 6;
            this.label2.Text = "ФИО студента:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.label3.Location = new System.Drawing.Point(352, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(149, 24);
            this.label3.TabIndex = 7;
            this.label3.Text = "Список группы:";
            // 
            // btn_upload_file
            // 
            this.btn_upload_file.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.btn_upload_file.Location = new System.Drawing.Point(12, 268);
            this.btn_upload_file.Name = "btn_upload_file";
            this.btn_upload_file.Size = new System.Drawing.Size(233, 39);
            this.btn_upload_file.TabIndex = 8;
            this.btn_upload_file.Text = "Загрузить из файла";
            this.btn_upload_file.UseVisualStyleBackColor = true;
            this.btn_upload_file.Click += new System.EventHandler(this.btn_upload_file_Click);
            // 
            // EditGroupForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(628, 350);
            this.Controls.Add(this.btn_upload_file);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button_cancel);
            this.Controls.Add(this.button_save);
            this.Controls.Add(this.button_remove);
            this.Controls.Add(this.button_add);
            this.Controls.Add(this.textBox_NewStudent);
            this.Controls.Add(this.listBox_Students);
            this.Controls.Add(this.comboBox_Groups);
            this.MinimumSize = new System.Drawing.Size(644, 389);
            this.Name = "EditGroupForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Редактирование группы";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBox_Groups;
        private System.Windows.Forms.ListBox listBox_Students;
        private System.Windows.Forms.TextBox textBox_NewStudent;
        private System.Windows.Forms.Button button_add;
        private System.Windows.Forms.Button button_remove;
        private System.Windows.Forms.Button button_save;
        private System.Windows.Forms.Button button_cancel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btn_upload_file;
    }
}