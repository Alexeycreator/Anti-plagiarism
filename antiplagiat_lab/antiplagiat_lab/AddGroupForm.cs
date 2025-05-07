using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
namespace antiplagiat_lab
{
    public partial class AddGroupForm : Form
    {
        private List<Group> groups;

        public AddGroupForm(List<Group> groups)
        {
            InitializeComponent();
            this.groups = groups;
        }
        #region Btn
        private void buttonAddStudent_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(textBox_NewStudent.Text))
            {
                listBox_Students.Items.Add(textBox_NewStudent.Text);
                textBox_NewStudent.Clear();
            }
        }
        private void buttonDeleteStudent_Click(object sender, EventArgs e)
        {
            if (listBox_Students.SelectedItem != null)
            {
                listBox_Students.Items.Remove(listBox_Students.SelectedItem);
            }
            else
            {
                MessageBox.Show("Выберите студента для удаления.");
            }
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(textBox_GroupName.Text))
            {
                string baseName = textBox_GroupName.Text;
                string newName = baseName;
                int counter = 1;

                while (groups.Any(g => g.Name == newName))
                {
                    newName = $"{baseName} ({counter})";
                    counter++;
                }

                var students = listBox_Students.Items.Cast<string>()
                                                     .Select(name => new Student { Name = name })
                                                     .ToList();

                groups.Add(new Group { Name = newName, Students = students });

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                MessageBox.Show("Пожалуйста, введите название группы.");
            }
        }


        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void btn_load_st_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Текстовые файлы (*.txt)|*.txt";
                openFileDialog.Multiselect = false;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;

                    try
                    {

                        string[] lines = File.ReadAllLines(filePath);


                        foreach (string line in lines)
                        {
                            listBox_Students.Items.Add(line);
                        }

                        MessageBox.Show("Файл успешно загружен!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при загрузке файла:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            
        }
    }
        #endregion


    }
}
