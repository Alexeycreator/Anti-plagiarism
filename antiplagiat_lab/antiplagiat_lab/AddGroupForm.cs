using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

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
    }
}
