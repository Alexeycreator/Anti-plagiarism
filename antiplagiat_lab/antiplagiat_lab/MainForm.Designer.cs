namespace antiplagiat_lab
{
    partial class MainForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
      this.panel_student = new System.Windows.Forms.Panel();
      this.numUD_NumberLab = new System.Windows.Forms.NumericUpDown();
      this.labelNumberLab = new System.Windows.Forms.Label();
      this.comboBox_Student = new System.Windows.Forms.ComboBox();
      this.comboBox_Group = new System.Windows.Forms.ComboBox();
      this.label_InfoStudent = new System.Windows.Forms.Label();
      this.label_Group = new System.Windows.Forms.Label();
      this.label_Student = new System.Windows.Forms.Label();
      this.panel_report = new System.Windows.Forms.Panel();
      this.panel_addCode = new System.Windows.Forms.Panel();
      this.label_addCode = new System.Windows.Forms.Label();
      this.label_filesCode = new System.Windows.Forms.Label();
      this.panel_newPeport = new System.Windows.Forms.Panel();
      this.label_newReport = new System.Windows.Forms.Label();
      this.label_Filesreport = new System.Windows.Forms.Label();
      this.progressBar = new System.Windows.Forms.ProgressBar();
      this.label_report = new System.Windows.Forms.Label();
      this.panel_infoReport = new System.Windows.Forms.Panel();
      this.label_SummASCII = new System.Windows.Forms.Label();
      this.comboBox_currentReport = new System.Windows.Forms.ComboBox();
      this.label_CountSymbol = new System.Windows.Forms.Label();
      this.label_CountWords = new System.Windows.Forms.Label();
      this.label_nameSummASCII = new System.Windows.Forms.Label();
      this.label_NameCountSymbol = new System.Windows.Forms.Label();
      this.label_NameCountWords = new System.Windows.Forms.Label();
      this.label_InfoReport = new System.Windows.Forms.Label();
      this.panel_Concidence = new System.Windows.Forms.Panel();
      this.dataGridView_Coincidence = new System.Windows.Forms.DataGridView();
      this.Student = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.Group = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.NumberLab = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.summASCII = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.percentMath = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.NameFile = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.PathFile = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.button_Open = new System.Windows.Forms.Button();
      this.label_Coincidence = new System.Windows.Forms.Label();
      this.menuStrip1 = new System.Windows.Forms.MenuStrip();
      this.ToolStripMenuItem_addGroup = new System.Windows.Forms.ToolStripMenuItem();
      this.ToolStripMenuItem_editGroup = new System.Windows.Forms.ToolStripMenuItem();
      this.ToolStripMenuItem_menuReport = new System.Windows.Forms.ToolStripMenuItem();
      this.panel1 = new System.Windows.Forms.Panel();
      this.label_countELSE = new System.Windows.Forms.Label();
      this.label_countIF = new System.Windows.Forms.Label();
      this.label_countWHILE = new System.Windows.Forms.Label();
      this.label_countFOR = new System.Windows.Forms.Label();
      this.label7 = new System.Windows.Forms.Label();
      this.label5 = new System.Windows.Forms.Label();
      this.label4 = new System.Windows.Forms.Label();
      this.label2 = new System.Windows.Forms.Label();
      this.label1 = new System.Windows.Forms.Label();
      this.verticalScrollBar = new System.Windows.Forms.VScrollBar();
      this.panel_student.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize)(this.numUD_NumberLab)).BeginInit();
      this.panel_report.SuspendLayout();
      this.panel_addCode.SuspendLayout();
      this.panel_newPeport.SuspendLayout();
      this.panel_infoReport.SuspendLayout();
      this.panel_Concidence.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize)(this.dataGridView_Coincidence)).BeginInit();
      this.menuStrip1.SuspendLayout();
      this.panel1.SuspendLayout();
      this.SuspendLayout();
      // 
      // panel_student
      // 
      this.panel_student.Controls.Add(this.numUD_NumberLab);
      this.panel_student.Controls.Add(this.labelNumberLab);
      this.panel_student.Controls.Add(this.comboBox_Student);
      this.panel_student.Controls.Add(this.comboBox_Group);
      this.panel_student.Controls.Add(this.label_InfoStudent);
      this.panel_student.Controls.Add(this.label_Group);
      this.panel_student.Controls.Add(this.label_Student);
      this.panel_student.Location = new System.Drawing.Point(16, 46);
      this.panel_student.Margin = new System.Windows.Forms.Padding(4);
      this.panel_student.Name = "panel_student";
      this.panel_student.Size = new System.Drawing.Size(413, 246);
      this.panel_student.TabIndex = 1;
      // 
      // numUD_NumberLab
      // 
      this.numUD_NumberLab.Font = new System.Drawing.Font("Times New Roman", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.numUD_NumberLab.Location = new System.Drawing.Point(151, 200);
      this.numUD_NumberLab.Name = "numUD_NumberLab";
      this.numUD_NumberLab.Size = new System.Drawing.Size(240, 34);
      this.numUD_NumberLab.TabIndex = 8;
      this.numUD_NumberLab.ValueChanged += new System.EventHandler(this.numUD_NumberLab_ValueChanged);
      this.numUD_NumberLab.Enter += new System.EventHandler(this.numUD_NumberLab_Enter);
      // 
      // labelNumberLab
      // 
      this.labelNumberLab.AutoSize = true;
      this.labelNumberLab.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.labelNumberLab.Location = new System.Drawing.Point(12, 205);
      this.labelNumberLab.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.labelNumberLab.Name = "labelNumberLab";
      this.labelNumberLab.Size = new System.Drawing.Size(132, 29);
      this.labelNumberLab.TabIndex = 7;
      this.labelNumberLab.Text = "Номер л/р";
      // 
      // comboBox_Student
      // 
      this.comboBox_Student.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
      this.comboBox_Student.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
      this.comboBox_Student.FormattingEnabled = true;
      this.comboBox_Student.Location = new System.Drawing.Point(12, 150);
      this.comboBox_Student.Margin = new System.Windows.Forms.Padding(4);
      this.comboBox_Student.Name = "comboBox_Student";
      this.comboBox_Student.Size = new System.Drawing.Size(379, 37);
      this.comboBox_Student.TabIndex = 6;
      this.comboBox_Student.SelectedIndexChanged += new System.EventHandler(this.comboBox_Student_SelectedIndexChanged);
      // 
      // comboBox_Group
      // 
      this.comboBox_Group.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
      this.comboBox_Group.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
      this.comboBox_Group.FormattingEnabled = true;
      this.comboBox_Group.Location = new System.Drawing.Point(12, 77);
      this.comboBox_Group.Margin = new System.Windows.Forms.Padding(4);
      this.comboBox_Group.Name = "comboBox_Group";
      this.comboBox_Group.Size = new System.Drawing.Size(233, 37);
      this.comboBox_Group.TabIndex = 6;
      this.comboBox_Group.SelectedIndexChanged += new System.EventHandler(this.comboBox_Group_SelectedIndexChanged);
      // 
      // label_InfoStudent
      // 
      this.label_InfoStudent.AutoSize = true;
      this.label_InfoStudent.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_InfoStudent.Location = new System.Drawing.Point(4, 12);
      this.label_InfoStudent.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_InfoStudent.Name = "label_InfoStudent";
      this.label_InfoStudent.Size = new System.Drawing.Size(325, 29);
      this.label_InfoStudent.TabIndex = 5;
      this.label_InfoStudent.Text = "Информация о студенте:";
      // 
      // label_Group
      // 
      this.label_Group.AutoSize = true;
      this.label_Group.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_Group.Location = new System.Drawing.Point(7, 43);
      this.label_Group.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_Group.Name = "label_Group";
      this.label_Group.Size = new System.Drawing.Size(101, 29);
      this.label_Group.TabIndex = 3;
      this.label_Group.Text = "Группа:";
      // 
      // label_Student
      // 
      this.label_Student.AutoSize = true;
      this.label_Student.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_Student.Location = new System.Drawing.Point(7, 117);
      this.label_Student.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_Student.Name = "label_Student";
      this.label_Student.Size = new System.Drawing.Size(114, 29);
      this.label_Student.TabIndex = 1;
      this.label_Student.Text = "Студент:";
      // 
      // panel_report
      // 
      this.panel_report.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.panel_report.Controls.Add(this.panel_addCode);
      this.panel_report.Controls.Add(this.panel_newPeport);
      this.panel_report.Controls.Add(this.progressBar);
      this.panel_report.Controls.Add(this.label_report);
      this.panel_report.Location = new System.Drawing.Point(437, 46);
      this.panel_report.Margin = new System.Windows.Forms.Padding(4);
      this.panel_report.Name = "panel_report";
      this.panel_report.Size = new System.Drawing.Size(653, 246);
      this.panel_report.TabIndex = 2;
      // 
      // panel_addCode
      // 
      this.panel_addCode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.panel_addCode.Controls.Add(this.label_addCode);
      this.panel_addCode.Controls.Add(this.label_filesCode);
      this.panel_addCode.Location = new System.Drawing.Point(333, 48);
      this.panel_addCode.Margin = new System.Windows.Forms.Padding(4);
      this.panel_addCode.Name = "panel_addCode";
      this.panel_addCode.Size = new System.Drawing.Size(306, 147);
      this.panel_addCode.TabIndex = 10;
      this.panel_addCode.Click += new System.EventHandler(this.addCode_Click);
      // 
      // label_addCode
      // 
      this.label_addCode.AutoSize = true;
      this.label_addCode.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_addCode.Location = new System.Drawing.Point(3, 5);
      this.label_addCode.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_addCode.Name = "label_addCode";
      this.label_addCode.Size = new System.Drawing.Size(169, 29);
      this.label_addCode.TabIndex = 6;
      this.label_addCode.Text = "Добавить код";
      this.label_addCode.Click += new System.EventHandler(this.addCode_Click);
      // 
      // label_filesCode
      // 
      this.label_filesCode.AutoSize = true;
      this.label_filesCode.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_filesCode.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
      this.label_filesCode.Location = new System.Drawing.Point(56, 63);
      this.label_filesCode.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_filesCode.Name = "label_filesCode";
      this.label_filesCode.Size = new System.Drawing.Size(201, 58);
      this.label_filesCode.TabIndex = 5;
      this.label_filesCode.Text = "нажмите, чтобы\r\nвыбрать файл";
      this.label_filesCode.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
      this.label_filesCode.Click += new System.EventHandler(this.addCode_Click);
      // 
      // panel_newPeport
      // 
      this.panel_newPeport.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.panel_newPeport.Controls.Add(this.label_newReport);
      this.panel_newPeport.Controls.Add(this.label_Filesreport);
      this.panel_newPeport.Location = new System.Drawing.Point(13, 48);
      this.panel_newPeport.Margin = new System.Windows.Forms.Padding(4);
      this.panel_newPeport.Name = "panel_newPeport";
      this.panel_newPeport.Size = new System.Drawing.Size(306, 147);
      this.panel_newPeport.TabIndex = 10;
      this.panel_newPeport.Click += new System.EventHandler(this.FilesReport_Click);
      // 
      // label_newReport
      // 
      this.label_newReport.AutoSize = true;
      this.label_newReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_newReport.Location = new System.Drawing.Point(3, 5);
      this.label_newReport.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_newReport.Name = "label_newReport";
      this.label_newReport.Size = new System.Drawing.Size(275, 29);
      this.label_newReport.TabIndex = 6;
      this.label_newReport.Text = "Добавить новый отчет";
      this.label_newReport.Click += new System.EventHandler(this.FilesReport_Click);
      // 
      // label_Filesreport
      // 
      this.label_Filesreport.AutoSize = true;
      this.label_Filesreport.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_Filesreport.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
      this.label_Filesreport.Location = new System.Drawing.Point(45, 63);
      this.label_Filesreport.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_Filesreport.Name = "label_Filesreport";
      this.label_Filesreport.Size = new System.Drawing.Size(201, 58);
      this.label_Filesreport.TabIndex = 5;
      this.label_Filesreport.Text = "нажмите, чтобы\r\nвыбрать файл";
      this.label_Filesreport.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
      this.label_Filesreport.Click += new System.EventHandler(this.FilesReport_Click);
      // 
      // progressBar
      // 
      this.progressBar.Location = new System.Drawing.Point(13, 204);
      this.progressBar.Margin = new System.Windows.Forms.Padding(4);
      this.progressBar.Name = "progressBar";
      this.progressBar.Size = new System.Drawing.Size(627, 31);
      this.progressBar.TabIndex = 9;
      // 
      // label_report
      // 
      this.label_report.AutoSize = true;
      this.label_report.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_report.Location = new System.Drawing.Point(17, 12);
      this.label_report.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_report.Name = "label_report";
      this.label_report.Size = new System.Drawing.Size(89, 29);
      this.label_report.TabIndex = 4;
      this.label_report.Text = "Отчет";
      // 
      // panel_infoReport
      // 
      this.panel_infoReport.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.panel_infoReport.Controls.Add(this.label_SummASCII);
      this.panel_infoReport.Controls.Add(this.comboBox_currentReport);
      this.panel_infoReport.Controls.Add(this.label_CountSymbol);
      this.panel_infoReport.Controls.Add(this.label_CountWords);
      this.panel_infoReport.Controls.Add(this.label_nameSummASCII);
      this.panel_infoReport.Controls.Add(this.label_NameCountSymbol);
      this.panel_infoReport.Controls.Add(this.label_NameCountWords);
      this.panel_infoReport.Controls.Add(this.label_InfoReport);
      this.panel_infoReport.Location = new System.Drawing.Point(16, 299);
      this.panel_infoReport.Margin = new System.Windows.Forms.Padding(4);
      this.panel_infoReport.Name = "panel_infoReport";
      this.panel_infoReport.Size = new System.Drawing.Size(747, 299);
      this.panel_infoReport.TabIndex = 6;
      // 
      // label_SummASCII
      // 
      this.label_SummASCII.AutoSize = true;
      this.label_SummASCII.BackColor = System.Drawing.SystemColors.Window;
      this.label_SummASCII.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.label_SummASCII.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_SummASCII.Location = new System.Drawing.Point(16, 234);
      this.label_SummASCII.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_SummASCII.Name = "label_SummASCII";
      this.label_SummASCII.Size = new System.Drawing.Size(57, 31);
      this.label_SummASCII.TabIndex = 10;
      this.label_SummASCII.Text = "       ";
      // 
      // comboBox_currentReport
      // 
      this.comboBox_currentReport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
      this.comboBox_currentReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
      this.comboBox_currentReport.FormattingEnabled = true;
      this.comboBox_currentReport.Location = new System.Drawing.Point(341, 12);
      this.comboBox_currentReport.Margin = new System.Windows.Forms.Padding(4);
      this.comboBox_currentReport.Name = "comboBox_currentReport";
      this.comboBox_currentReport.Size = new System.Drawing.Size(381, 37);
      this.comboBox_currentReport.TabIndex = 6;
      this.comboBox_currentReport.SelectedIndexChanged += new System.EventHandler(this.comboBox_currentReport_SelectedIndexChanged);
      // 
      // label_CountSymbol
      // 
      this.label_CountSymbol.AutoSize = true;
      this.label_CountSymbol.BackColor = System.Drawing.SystemColors.Window;
      this.label_CountSymbol.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.label_CountSymbol.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_CountSymbol.Location = new System.Drawing.Point(16, 160);
      this.label_CountSymbol.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_CountSymbol.Name = "label_CountSymbol";
      this.label_CountSymbol.Size = new System.Drawing.Size(57, 31);
      this.label_CountSymbol.TabIndex = 9;
      this.label_CountSymbol.Text = "       ";
      // 
      // label_CountWords
      // 
      this.label_CountWords.AutoSize = true;
      this.label_CountWords.BackColor = System.Drawing.SystemColors.Window;
      this.label_CountWords.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.label_CountWords.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_CountWords.Location = new System.Drawing.Point(16, 86);
      this.label_CountWords.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_CountWords.Name = "label_CountWords";
      this.label_CountWords.Size = new System.Drawing.Size(57, 31);
      this.label_CountWords.TabIndex = 8;
      this.label_CountWords.Text = "       ";
      // 
      // label_nameSummASCII
      // 
      this.label_nameSummASCII.AutoSize = true;
      this.label_nameSummASCII.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_nameSummASCII.Location = new System.Drawing.Point(8, 203);
      this.label_nameSummASCII.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_nameSummASCII.Name = "label_nameSummASCII";
      this.label_nameSummASCII.Size = new System.Drawing.Size(357, 29);
      this.label_nameSummASCII.TabIndex = 7;
      this.label_nameSummASCII.Text = "Сумма ASCII-кодов символов:";
      // 
      // label_NameCountSymbol
      // 
      this.label_NameCountSymbol.AutoSize = true;
      this.label_NameCountSymbol.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_NameCountSymbol.Location = new System.Drawing.Point(8, 129);
      this.label_NameCountSymbol.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_NameCountSymbol.Name = "label_NameCountSymbol";
      this.label_NameCountSymbol.Size = new System.Drawing.Size(278, 29);
      this.label_NameCountSymbol.TabIndex = 6;
      this.label_NameCountSymbol.Text = "Количество символов:";
      // 
      // label_NameCountWords
      // 
      this.label_NameCountWords.AutoSize = true;
      this.label_NameCountWords.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_NameCountWords.Location = new System.Drawing.Point(8, 55);
      this.label_NameCountWords.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_NameCountWords.Name = "label_NameCountWords";
      this.label_NameCountWords.Size = new System.Drawing.Size(218, 29);
      this.label_NameCountWords.TabIndex = 6;
      this.label_NameCountWords.Text = "Количество слов:";
      // 
      // label_InfoReport
      // 
      this.label_InfoReport.AutoSize = true;
      this.label_InfoReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_InfoReport.Location = new System.Drawing.Point(5, 16);
      this.label_InfoReport.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_InfoReport.Name = "label_InfoReport";
      this.label_InfoReport.Size = new System.Drawing.Size(314, 29);
      this.label_InfoReport.TabIndex = 4;
      this.label_InfoReport.Text = "Информация об отчете:";
      // 
      // panel_Concidence
      // 
      this.panel_Concidence.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.panel_Concidence.Controls.Add(this.dataGridView_Coincidence);
      this.panel_Concidence.Controls.Add(this.button_Open);
      this.panel_Concidence.Controls.Add(this.label_Coincidence);
      this.panel_Concidence.Location = new System.Drawing.Point(16, 606);
      this.panel_Concidence.Margin = new System.Windows.Forms.Padding(4);
      this.panel_Concidence.Name = "panel_Concidence";
      this.panel_Concidence.Size = new System.Drawing.Size(1074, 381);
      this.panel_Concidence.TabIndex = 7;
      // 
      // dataGridView_Coincidence
      // 
      this.dataGridView_Coincidence.AllowUserToAddRows = false;
      this.dataGridView_Coincidence.AllowUserToDeleteRows = false;
      this.dataGridView_Coincidence.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
      this.dataGridView_Coincidence.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Student,
            this.Group,
            this.NumberLab,
            this.summASCII,
            this.percentMath,
            this.NameFile,
            this.PathFile});
      this.dataGridView_Coincidence.Location = new System.Drawing.Point(9, 46);
      this.dataGridView_Coincidence.Margin = new System.Windows.Forms.Padding(4);
      this.dataGridView_Coincidence.Name = "dataGridView_Coincidence";
      this.dataGridView_Coincidence.ReadOnly = true;
      this.dataGridView_Coincidence.RowHeadersWidth = 51;
      this.dataGridView_Coincidence.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
      this.dataGridView_Coincidence.Size = new System.Drawing.Size(1053, 271);
      this.dataGridView_Coincidence.TabIndex = 7;
      // 
      // Student
      // 
      this.Student.HeaderText = "Студент";
      this.Student.MinimumWidth = 6;
      this.Student.Name = "Student";
      this.Student.ReadOnly = true;
      this.Student.Width = 125;
      // 
      // Group
      // 
      this.Group.HeaderText = "Группа";
      this.Group.MinimumWidth = 6;
      this.Group.Name = "Group";
      this.Group.ReadOnly = true;
      this.Group.Width = 125;
      // 
      // NumberLab
      // 
      this.NumberLab.HeaderText = "Номер работы";
      this.NumberLab.MinimumWidth = 6;
      this.NumberLab.Name = "NumberLab";
      this.NumberLab.ReadOnly = true;
      this.NumberLab.Width = 125;
      // 
      // summASCII
      // 
      this.summASCII.HeaderText = "Сумма ASCII-кода";
      this.summASCII.MinimumWidth = 6;
      this.summASCII.Name = "summASCII";
      this.summASCII.ReadOnly = true;
      this.summASCII.Width = 125;
      // 
      // percentMath
      // 
      this.percentMath.HeaderText = "Процент совпадения";
      this.percentMath.MinimumWidth = 6;
      this.percentMath.Name = "percentMath";
      this.percentMath.ReadOnly = true;
      this.percentMath.Width = 125;
      // 
      // NameFile
      // 
      this.NameFile.HeaderText = "Наименование файла";
      this.NameFile.MinimumWidth = 6;
      this.NameFile.Name = "NameFile";
      this.NameFile.ReadOnly = true;
      this.NameFile.Width = 125;
      // 
      // PathFile
      // 
      this.PathFile.HeaderText = "Путь к файлу";
      this.PathFile.MinimumWidth = 6;
      this.PathFile.Name = "PathFile";
      this.PathFile.ReadOnly = true;
      this.PathFile.Width = 125;
      // 
      // button_Open
      // 
      this.button_Open.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.button_Open.Location = new System.Drawing.Point(888, 324);
      this.button_Open.Margin = new System.Windows.Forms.Padding(4);
      this.button_Open.Name = "button_Open";
      this.button_Open.Size = new System.Drawing.Size(167, 39);
      this.button_Open.TabIndex = 6;
      this.button_Open.Text = "Открыть";
      this.button_Open.UseVisualStyleBackColor = true;
      this.button_Open.Click += new System.EventHandler(this.button_Open_Click);
      // 
      // label_Coincidence
      // 
      this.label_Coincidence.AutoSize = true;
      this.label_Coincidence.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_Coincidence.Location = new System.Drawing.Point(17, 12);
      this.label_Coincidence.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_Coincidence.Name = "label_Coincidence";
      this.label_Coincidence.Size = new System.Drawing.Size(174, 29);
      this.label_Coincidence.TabIndex = 4;
      this.label_Coincidence.Text = "Совпадения:";
      // 
      // menuStrip1
      // 
      this.menuStrip1.BackColor = System.Drawing.Color.Silver;
      this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
      this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ToolStripMenuItem_addGroup,
            this.ToolStripMenuItem_editGroup,
            this.ToolStripMenuItem_menuReport});
      this.menuStrip1.Location = new System.Drawing.Point(0, 0);
      this.menuStrip1.Name = "menuStrip1";
      this.menuStrip1.Size = new System.Drawing.Size(1125, 36);
      this.menuStrip1.TabIndex = 8;
      this.menuStrip1.Text = "menuStrip1";
      // 
      // ToolStripMenuItem_addGroup
      // 
      this.ToolStripMenuItem_addGroup.BackColor = System.Drawing.SystemColors.Control;
      this.ToolStripMenuItem_addGroup.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.ToolStripMenuItem_addGroup.Name = "ToolStripMenuItem_addGroup";
      this.ToolStripMenuItem_addGroup.Size = new System.Drawing.Size(188, 32);
      this.ToolStripMenuItem_addGroup.Text = "Добавить группу";
      this.ToolStripMenuItem_addGroup.Click += new System.EventHandler(this.ToolStripMenuItem_addGroup_Click);
      // 
      // ToolStripMenuItem_editGroup
      // 
      this.ToolStripMenuItem_editGroup.BackColor = System.Drawing.SystemColors.Control;
      this.ToolStripMenuItem_editGroup.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.ToolStripMenuItem_editGroup.Name = "ToolStripMenuItem_editGroup";
      this.ToolStripMenuItem_editGroup.Size = new System.Drawing.Size(237, 32);
      this.ToolStripMenuItem_editGroup.Text = "Редактировать группу";
      this.ToolStripMenuItem_editGroup.Click += new System.EventHandler(this.ToolStripMenuItem_editGroup_Click);
      // 
      // ToolStripMenuItem_menuReport
      // 
      this.ToolStripMenuItem_menuReport.BackColor = System.Drawing.SystemColors.Control;
      this.ToolStripMenuItem_menuReport.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.ToolStripMenuItem_menuReport.Name = "ToolStripMenuItem_menuReport";
      this.ToolStripMenuItem_menuReport.Size = new System.Drawing.Size(235, 32);
      this.ToolStripMenuItem_menuReport.Text = "Управление отчётами";
      this.ToolStripMenuItem_menuReport.Click += new System.EventHandler(this.ToolStripMenuItem_menuReport_Click);
      // 
      // panel1
      // 
      this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.panel1.Controls.Add(this.label_countELSE);
      this.panel1.Controls.Add(this.label_countIF);
      this.panel1.Controls.Add(this.label_countWHILE);
      this.panel1.Controls.Add(this.label_countFOR);
      this.panel1.Controls.Add(this.label7);
      this.panel1.Controls.Add(this.label5);
      this.panel1.Controls.Add(this.label4);
      this.panel1.Controls.Add(this.label2);
      this.panel1.Controls.Add(this.label1);
      this.panel1.Location = new System.Drawing.Point(772, 299);
      this.panel1.Margin = new System.Windows.Forms.Padding(4);
      this.panel1.Name = "panel1";
      this.panel1.Size = new System.Drawing.Size(318, 299);
      this.panel1.TabIndex = 6;
      // 
      // label_countELSE
      // 
      this.label_countELSE.AutoSize = true;
      this.label_countELSE.BackColor = System.Drawing.SystemColors.Window;
      this.label_countELSE.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.label_countELSE.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_countELSE.Location = new System.Drawing.Point(112, 167);
      this.label_countELSE.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_countELSE.Name = "label_countELSE";
      this.label_countELSE.Size = new System.Drawing.Size(57, 31);
      this.label_countELSE.TabIndex = 10;
      this.label_countELSE.Text = "       ";
      // 
      // label_countIF
      // 
      this.label_countIF.AutoSize = true;
      this.label_countIF.BackColor = System.Drawing.SystemColors.Window;
      this.label_countIF.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.label_countIF.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_countIF.Location = new System.Drawing.Point(112, 130);
      this.label_countIF.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_countIF.Name = "label_countIF";
      this.label_countIF.Size = new System.Drawing.Size(57, 31);
      this.label_countIF.TabIndex = 10;
      this.label_countIF.Text = "       ";
      // 
      // label_countWHILE
      // 
      this.label_countWHILE.AutoSize = true;
      this.label_countWHILE.BackColor = System.Drawing.SystemColors.Window;
      this.label_countWHILE.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.label_countWHILE.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_countWHILE.Location = new System.Drawing.Point(112, 94);
      this.label_countWHILE.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_countWHILE.Name = "label_countWHILE";
      this.label_countWHILE.Size = new System.Drawing.Size(57, 31);
      this.label_countWHILE.TabIndex = 9;
      this.label_countWHILE.Text = "       ";
      // 
      // label_countFOR
      // 
      this.label_countFOR.AutoSize = true;
      this.label_countFOR.BackColor = System.Drawing.SystemColors.Window;
      this.label_countFOR.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.label_countFOR.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label_countFOR.Location = new System.Drawing.Point(112, 57);
      this.label_countFOR.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label_countFOR.Name = "label_countFOR";
      this.label_countFOR.Size = new System.Drawing.Size(57, 31);
      this.label_countFOR.TabIndex = 8;
      this.label_countFOR.Text = "       ";
      // 
      // label7
      // 
      this.label7.AutoSize = true;
      this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label7.Location = new System.Drawing.Point(5, 16);
      this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label7.Name = "label7";
      this.label7.Size = new System.Drawing.Size(271, 29);
      this.label7.TabIndex = 4;
      this.label7.Text = "Информация о коде:";
      // 
      // label5
      // 
      this.label5.AutoSize = true;
      this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label5.Location = new System.Drawing.Point(20, 167);
      this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label5.Name = "label5";
      this.label5.Size = new System.Drawing.Size(80, 29);
      this.label5.TabIndex = 6;
      this.label5.Text = "ELSE:";
      // 
      // label4
      // 
      this.label4.AutoSize = true;
      this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label4.Location = new System.Drawing.Point(63, 130);
      this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label4.Name = "label4";
      this.label4.Size = new System.Drawing.Size(40, 29);
      this.label4.TabIndex = 6;
      this.label4.Text = "IF:";
      // 
      // label2
      // 
      this.label2.AutoSize = true;
      this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label2.Location = new System.Drawing.Point(5, 94);
      this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label2.Name = "label2";
      this.label2.Size = new System.Drawing.Size(93, 29);
      this.label2.TabIndex = 6;
      this.label2.Text = "WHILE:";
      // 
      // label1
      // 
      this.label1.AutoSize = true;
      this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      this.label1.Location = new System.Drawing.Point(31, 57);
      this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
      this.label1.Name = "label1";
      this.label1.Size = new System.Drawing.Size(70, 29);
      this.label1.TabIndex = 6;
      this.label1.Text = "FOR:";
      // 
      // verticalScrollBar
      // 
      this.verticalScrollBar.Location = new System.Drawing.Point(1094, 36);
      this.verticalScrollBar.Name = "verticalScrollBar";
      this.verticalScrollBar.Size = new System.Drawing.Size(25, 951);
      this.verticalScrollBar.TabIndex = 9;
      // 
      // MainForm
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(1125, 998);
      this.Controls.Add(this.verticalScrollBar);
      this.Controls.Add(this.panel_Concidence);
      this.Controls.Add(this.panel1);
      this.Controls.Add(this.panel_infoReport);
      this.Controls.Add(this.panel_report);
      this.Controls.Add(this.panel_student);
      this.Controls.Add(this.menuStrip1);
      this.MainMenuStrip = this.menuStrip1;
      this.Margin = new System.Windows.Forms.Padding(4);
      this.Name = "MainForm";
      this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
      this.Text = "Антиплагиат";
      this.Load += new System.EventHandler(this.MainForm_Load);
      this.panel_student.ResumeLayout(false);
      this.panel_student.PerformLayout();
      ((System.ComponentModel.ISupportInitialize)(this.numUD_NumberLab)).EndInit();
      this.panel_report.ResumeLayout(false);
      this.panel_report.PerformLayout();
      this.panel_addCode.ResumeLayout(false);
      this.panel_addCode.PerformLayout();
      this.panel_newPeport.ResumeLayout(false);
      this.panel_newPeport.PerformLayout();
      this.panel_infoReport.ResumeLayout(false);
      this.panel_infoReport.PerformLayout();
      this.panel_Concidence.ResumeLayout(false);
      this.panel_Concidence.PerformLayout();
      ((System.ComponentModel.ISupportInitialize)(this.dataGridView_Coincidence)).EndInit();
      this.menuStrip1.ResumeLayout(false);
      this.menuStrip1.PerformLayout();
      this.panel1.ResumeLayout(false);
      this.panel1.PerformLayout();
      this.ResumeLayout(false);
      this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Panel panel_student;
        private System.Windows.Forms.Label label_Student;
        private System.Windows.Forms.Label label_Group;
        private System.Windows.Forms.Panel panel_report;
        private System.Windows.Forms.Label label_Filesreport;
        private System.Windows.Forms.Label label_report;
        private System.Windows.Forms.Label label_InfoStudent;
        private System.Windows.Forms.Panel panel_infoReport;
        private System.Windows.Forms.Label label_NameCountWords;
        private System.Windows.Forms.Label label_InfoReport;
        private System.Windows.Forms.Panel panel_Concidence;
        private System.Windows.Forms.Label label_Coincidence;
        private System.Windows.Forms.Label label_SummASCII;
        private System.Windows.Forms.Label label_CountSymbol;
        private System.Windows.Forms.Label label_CountWords;
        private System.Windows.Forms.Label label_nameSummASCII;
        private System.Windows.Forms.Label label_NameCountSymbol;
        private System.Windows.Forms.Button button_Open;
        private System.Windows.Forms.DataGridView dataGridView_Coincidence;
        private System.Windows.Forms.ComboBox comboBox_Group;
        private System.Windows.Forms.ComboBox comboBox_Student;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItem_addGroup;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItem_editGroup;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItem_menuReport;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Panel panel_newPeport;
        private System.Windows.Forms.ComboBox comboBox_currentReport;
        private System.Windows.Forms.Label label_newReport;
        private System.Windows.Forms.Panel panel_addCode;
        private System.Windows.Forms.Label label_addCode;
        private System.Windows.Forms.Label label_filesCode;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label_countWHILE;
        private System.Windows.Forms.Label label_countFOR;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label_countELSE;
        private System.Windows.Forms.Label label_countIF;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
    private System.Windows.Forms.NumericUpDown numUD_NumberLab;
    private System.Windows.Forms.Label labelNumberLab;
    private System.Windows.Forms.DataGridViewTextBoxColumn Student;
    private System.Windows.Forms.DataGridViewTextBoxColumn Group;
    private System.Windows.Forms.DataGridViewTextBoxColumn NumberLab;
    private System.Windows.Forms.DataGridViewTextBoxColumn summASCII;
    private System.Windows.Forms.DataGridViewTextBoxColumn percentMath;
    private System.Windows.Forms.DataGridViewTextBoxColumn NameFile;
    private System.Windows.Forms.DataGridViewTextBoxColumn PathFile;
    private System.Windows.Forms.VScrollBar verticalScrollBar;
  }
}

