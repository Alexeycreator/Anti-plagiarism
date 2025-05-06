using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace antiplagiat_lab
{
    public partial class ReportForm : Form
    {
        private Student student;

        public ReportForm(Student student)
        {
            InitializeComponent();
            this.student = student;

            FillReportList();
        }

        private void FillReportList()
        {
            listBox_Reports.Items.Clear();
            foreach (var report in student.Reports)
            {
                listBox_Reports.Items.Add(report.FileName);
            }
        }

        private void buttonDeleteReport_Click(object sender, EventArgs e)
        {
            if (listBox_Reports.SelectedItem != null)
            {
                string reportName = listBox_Reports.SelectedItem.ToString();
                var report = student.Reports.FirstOrDefault(r => r.FileName == reportName);

                if (report != null)
                {
                    try
                    {
                        File.Delete(report.FilePath);
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show($"Ошибка при удалении файла: {ex.Message}");
                    }

                    student.Reports.Remove(report);
                    FillReportList();
                }
            }
            else
            {
                MessageBox.Show("Пожалуйста, выберите отчёт для удаления.");
            }
        }
    }
}
