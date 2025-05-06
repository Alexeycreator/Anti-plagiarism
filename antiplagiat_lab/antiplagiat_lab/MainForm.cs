using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;

namespace antiplagiat_lab
{
  public partial class MainForm : Form
  {
    private const string DataFilePath = "data.json";
    private const string ReportsDirectory = "Отчёты";
    private List<Group> groups = new List<Group>();
        //asd12
    public MainForm()
    {
      InitializeComponent();
      InitializeApp();
    }

    private void InitializeApp()
    {
      if (!File.Exists(DataFilePath))
      {
        File.WriteAllText(DataFilePath, "[]");
      }

      if (!Directory.Exists(ReportsDirectory))
      {
        Directory.CreateDirectory(ReportsDirectory);
      }

      progressBar.Visible = false;
      LoadData();
      FillComboBoxes();
    }

    private void LoadData()
    {
      string jsonData = File.ReadAllText(DataFilePath);
      if (!string.IsNullOrWhiteSpace(jsonData))
      {
        groups = JsonSerializer.Deserialize<List<Group>>(jsonData) ?? new List<Group>();
      }
    }

    private void SaveData()
    {
      string jsonData = JsonSerializer.Serialize(groups);
      File.WriteAllText(DataFilePath, jsonData);
    }

    private void FillComboBoxes()
    {
      comboBox_Group.Items.Clear();
      foreach (var group in groups)
      {
        comboBox_Group.Items.Add(group.Name);
      }
    }

    private void ToolStripMenuItem_addGroup_Click(object sender, EventArgs e)
    {
      using (var addGroupForm = new AddGroupForm(groups))
      {
        if (addGroupForm.ShowDialog() == DialogResult.OK)
        {
          SaveData();
          FillComboBoxes();
        }
      }
    }

    private void ToolStripMenuItem_editGroup_Click(object sender, EventArgs e)
    {
      using (var editGroupForm = new EditGroupForm(groups))
      {
        if (editGroupForm.ShowDialog() == DialogResult.OK)
        {
          SaveData();
          FillComboBoxes();
          UpdateStudentList();
        }
      }
    }
    private void ToolStripMenuItem_menuReport_Click(object sender, EventArgs e)
    {
      if (comboBox_Student.SelectedItem != null)
      {
        string studentName = comboBox_Student.SelectedItem.ToString();
        var selectedGroup = comboBox_Group.SelectedItem.ToString();
        var group = groups.FirstOrDefault(g => g.Name == selectedGroup);
        var student = group?.Students.FirstOrDefault(s => s.Name == studentName);

        if (student != null)
        {
          using (var reportForm = new ReportForm(student))
          {
            if (reportForm.ShowDialog() == DialogResult.OK)
            {
              SaveData();
              UpdateReportList();
            }
          }
        }
      }
      else
      {
        MessageBox.Show("Выберите студента для просмотра отчётов.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
      }
    }

    private void comboBox_Group_SelectedIndexChanged(object sender, EventArgs e)
    {
      UpdateStudentList();
      comboBox_Student.SelectedItem = null;
      comboBox_currentReport.SelectedItem = null;
      comboBox_Student.Enabled = true;
    }

    private void FilesReport_Click(object sender, EventArgs e)
    {
      if (comboBox_Group.SelectedItem != null && comboBox_Student.SelectedItem != null)
      {
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
          if (openFileDialog.ShowDialog() == DialogResult.OK)
          {
            string groupName = comboBox_Group.SelectedItem.ToString();
            string studentName = comboBox_Student.SelectedItem.ToString();
            var group = groups.FirstOrDefault(g => g.Name == groupName);
            var student = group?.Students.FirstOrDefault(s => s.Name == studentName);

            if (group != null && student != null)
            {
              string destPath = Path.Combine(ReportsDirectory, groupName, studentName);
              Directory.CreateDirectory(destPath);
              string destFile = Path.Combine(Environment.CurrentDirectory, destPath, Path.GetFileName(openFileDialog.FileName));

              bool fileExists = student.Reports.Any(r => r.FileName == Path.GetFileName(openFileDialog.FileName));

              if (fileExists)
              {
                var existingReport = student.Reports.First(r => r.FileName == Path.GetFileName(openFileDialog.FileName));

                var dialogResult = MessageBox.Show(
                    $"Отчет с таким названием уже существует.\n\n" +
                    $"Текущий файл: {existingReport.FileName}\n" +
                    $"Размер: {existingReport.SymbolCount} символов, {existingReport.WordCount} слов\n" +
                    $"Хотите заменить его новым файлом?\n\n" +
                    $"Новый файл: {Path.GetFileName(openFileDialog.FileName)}",
                    "Заменить файл?",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (dialogResult == DialogResult.Yes)
                {
                  student.Reports.Remove(existingReport);
                  File.Copy(openFileDialog.FileName, destFile, true);
                  var newReportData = AnalyzeReport(destFile);
                  student.Reports.Add(newReportData);

                  SaveData();
                  FillDataGridView(newReportData.AsciiSum, destFile);
                }
              }
              else
              {
                File.Copy(openFileDialog.FileName, destFile, true);
                var newReportData = AnalyzeReport(destFile);
                student.Reports.Add(newReportData);

                SaveData();
                comboBox_currentReport.Items.Clear();
                foreach (var report in student.Reports)
                {
                  comboBox_currentReport.Items.Add(report.FileName);
                }
                comboBox_currentReport.SelectedItem = newReportData.FileName;
              }
            }
          }
        }
      }
      else
      {
        MessageBox.Show("Пожалуйста, выберите группу и студента.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
      }
    }

    private ReportData AnalyzeReport(string filePath)
    {
      if (!File.Exists(filePath) || Path.GetExtension(filePath) != ".docx")
      {
        File.Delete(filePath);
        throw new FileNotFoundException("Файл не найден или формат файла не поддерживается.");
      }

      Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
      Document doc = null;

      try
      {
        doc = wordApp.Documents.Open(filePath, ReadOnly: true);
        int wordCount = doc.Words.Count - 1;
        int symbolCount = doc.Content.Characters.Count - 1;

        long asciiSum = 0;
        progressBar.Visible = true;
        progressBar.Maximum = doc.Content.Characters.Count;
        progressBar.Value = 0;

        foreach (Range character in doc.Content.Characters)
        {
          char currentChar = character.Text[0];
          int asciiCode = (int)currentChar;

          if (asciiCode >= 0 && asciiCode <= 127)
          {
            asciiSum += asciiCode;
          }
          else
          {
            byte[] tmp = Encoding.GetEncoding(1251).GetBytes(character.Text);
            foreach (var b in tmp)
            {
              asciiSum += b;
            }
          }
          progressBar.Value++;
        }

        label_CountWords.Text = wordCount.ToString();
        label_CountSymbol.Text = symbolCount.ToString();
        label_SummASCII.Text = asciiSum.ToString();

        return new ReportData
        {
          FileName = Path.GetFileName(filePath),
          FilePath = filePath,
          WordCount = wordCount,
          SymbolCount = symbolCount,
          AsciiSum = asciiSum
        };
      }
      finally
      {
        progressBar.Visible = false;
        doc?.Close(false);
        wordApp.Quit(false);
      }
    }
    private void addCode_Click(object sender, EventArgs e)
    {
      if (comboBox_Group.SelectedItem == null || comboBox_Student.SelectedItem == null || comboBox_currentReport.SelectedItem == null)
      {
        MessageBox.Show("Выберите группу, студента и отчет для добавления кода.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
      }

      string selectedGroup = comboBox_Group.SelectedItem.ToString();
      string selectedStudent = comboBox_Student.SelectedItem.ToString();
      string selectedReport = comboBox_currentReport.SelectedItem.ToString();

      var group = groups.FirstOrDefault(g => g.Name == selectedGroup);
      var student = group?.Students.FirstOrDefault(s => s.Name == selectedStudent);
      var report = student?.Reports.FirstOrDefault(r => r.FileName == selectedReport);

      if (report == null)
      {
        MessageBox.Show("Выбранный отчет не найден.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        return;
      }

      using (OpenFileDialog openFileDialog = new OpenFileDialog())
      {
        openFileDialog.Filter = "Text Files (*.txt)|*.txt";
        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
          if (report.CodeInfo != null)
          {
            var dialogResult = MessageBox.Show(
                "К отчету уже привязан код. Хотите заменить его новым?",
                "Код уже существует",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (dialogResult == DialogResult.No)
            {
              return;
            }
          }

          string code = File.ReadAllText(openFileDialog.FileName);
          report.CodeInfo = AnalyzeCode(code);

          SaveData();
          DisplayCodeInfo(report.CodeInfo);
        }
      }
    }

    private CodeAnalysis AnalyzeCode(string code)
    {
      return new CodeAnalysis
      {
        CountFor = CountOccurrences(code, @"\bfor\b"),
        CountWhile = CountOccurrences(code, @"\bwhile\b"),
        CountIf = CountOccurrences(code, @"\bif\b"),
        CountElse = CountOccurrences(code, @"\belse\b")
      };
    }

    private void DisplayCodeInfo(CodeAnalysis codeInfo)
    {
      if (codeInfo == null)
      {
        label_countFOR.Text = "       ";
        label_countWHILE.Text = "       ";
        label_countIF.Text = "       ";
        label_countELSE.Text = "       ";
      }
      else
      {
        label_countFOR.Text = $"{codeInfo.CountFor}";
        label_countWHILE.Text = $"{codeInfo.CountWhile}";
        label_countIF.Text = $"{codeInfo.CountIf}";
        label_countELSE.Text = $"{codeInfo.CountElse}";
      }
    }

    private int CountOccurrences(string text, string pattern)
    {
      return System.Text.RegularExpressions.Regex.Matches(text, pattern).Count;
    }

    private void FillDataGridView(long currentAsciiSum, string currentFilePath)
    {
      dataGridView_Coincidence.Rows.Clear();

      foreach (var group in groups)
      {
        foreach (var student in group.Students)
        {
          foreach (var report in student.Reports)
          {
            if (report.FilePath == currentFilePath)
              continue;
            double similarityPercentage = CalculateSimilarityPercentage(currentAsciiSum, report.AsciiSum);

            if (similarityPercentage <= 25)
            {
              dataGridView_Coincidence.Rows.Add(student.Name, group.Name, report.AsciiSum, Math.Round(100 - similarityPercentage, 2), report.FileName, report.FilePath, similarityPercentage.ToString("F2"));
            }
          }
        }
      }
    }

    private double CalculateSimilarityPercentage(long currentAsciiSum, long reportAsciiSum)
    {
      double maxAsciiSum = Math.Max(currentAsciiSum, reportAsciiSum);
      double percentage = (Math.Abs(currentAsciiSum - reportAsciiSum) / maxAsciiSum) * 100;

      return percentage;
    }

    private void button_Open_Click(object sender, EventArgs e)
    {
      if (dataGridView_Coincidence.SelectedRows.Count > 0)
      {
        string filePath = dataGridView_Coincidence.SelectedRows[0].Cells[5].Value.ToString();
        if (File.Exists(filePath))
        {
          System.Diagnostics.Process.Start(filePath);
        }
        else
        {
          MessageBox.Show("Файл не найден.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
      }
    }

    private void comboBox_Student_SelectedIndexChanged(object sender, EventArgs e)
    {
      UpdateReportList();
      comboBox_currentReport.SelectedItem = null;
      numUD_NumberLab.Enabled = true;
      SetNumericUpDownEnabled(true);
    }


    private void comboBox_currentReport_SelectedIndexChanged(object sender, EventArgs e)
    {
      if (comboBox_currentReport.SelectedItem == null)
      {
        label_CountWords.Text = "       ";
        label_CountSymbol.Text = "       ";
        label_SummASCII.Text = "       ";
        dataGridView_Coincidence.Rows.Clear();
        DisplayCodeInfo(null);
        return;
      }

      var selectedGroup = comboBox_Group.SelectedItem?.ToString();
      var selectedStudent = comboBox_Student.SelectedItem?.ToString();
      var selectedReport = comboBox_currentReport.SelectedItem?.ToString();

      if (selectedGroup == null || selectedStudent == null || selectedReport == null)
      {
        return;
      }

      var group = groups.FirstOrDefault(g => g.Name == selectedGroup);
      var student = group?.Students.FirstOrDefault(s => s.Name == selectedStudent);
      var report = student?.Reports.FirstOrDefault(r => r.FileName == selectedReport);

      if (report != null)
      {
        label_CountWords.Text = $"{report.WordCount}";
        label_CountSymbol.Text = $"{report.SymbolCount}";
        label_SummASCII.Text = $"{report.AsciiSum}";
        DisplayCodeInfo(report.CodeInfo);
        FillDataGridView(report.AsciiSum, report.FilePath);
      }
    }
    private void UpdateStudentList()
    {
      string selectedGroup = comboBox_Group.SelectedItem?.ToString();
      comboBox_Student.Items.Clear();
      if (selectedGroup != null)
      {
        var group = groups.FirstOrDefault(g => g.Name == selectedGroup);
        if (group != null)
        {
          foreach (var student in group.Students)
          {
            comboBox_Student.Items.Add(student.Name);
          }
        }
      }
    }
    private void UpdateReportList()
    {
      string selectedStudent = comboBox_Student.SelectedItem?.ToString();
      comboBox_currentReport.Items.Clear();
      if (selectedStudent != null)
      {
        var selectedGroup = comboBox_Group.SelectedItem?.ToString();
        var group = groups.FirstOrDefault(g => g.Name == selectedGroup);
        var student = group?.Students.FirstOrDefault(s => s.Name == selectedStudent);

        if (student != null && student.Reports.Any())
        {
          foreach (var report in student.Reports)
          {
            comboBox_currentReport.Items.Add(report.FileName);
          }
        }
      }
    }

    private void MainForm_Load(object sender, EventArgs e)
    {
      SettingsForm();
    }

    private void SettingsForm()
    {
      panel_newPeport.Click -= FilesReport_Click;
      label_Filesreport.Click -= FilesReport_Click;
      panel_addCode.Click -= addCode_Click;
      label_filesCode.Click -= addCode_Click;
      comboBox_Student.Enabled = false;
      numUD_NumberLab.Enabled = false;
      numUD_NumberLab.Value = 1;
      numUD_NumberLab.Maximum = 15;
      numUD_NumberLab.ReadOnly = true;
      numUD_NumberLab.TextAlign = HorizontalAlignment.Center;
    }

    private void CheckNumberLabs()
    {
      try
      {
        if (!numUD_NumberLab.Enabled)
        {
          ResetLabelsState();
          return;
        }
        if (numUD_NumberLab.Value == 0)
        {
          numUD_NumberLab.Value = 1;
          throw new FormatException($"Лабораторная работа с номером 0 не может существовать.");
        }
        else if (numUD_NumberLab.Value >= 1 && numUD_NumberLab.Enabled)
        {
          ActivateLabelsState();
        }
      }
      catch (FormatException fex)
      {
        MessageBox.Show($"{fex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        CheckErrorNumeric();
      }
      catch (Exception ex)
      {
        
        MessageBox.Show($"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        CheckErrorNumeric();
      }
    }

    private void ResetLabelsState()
    {
      label_Filesreport.Click -= FilesReport_Click;
      label_Filesreport.Visible = false;
      label_Filesreport.Enabled = false;
      label_filesCode.Click -= addCode_Click;
      label_filesCode.Visible = false;
      label_filesCode.Enabled = false;
    }

    private void ActivateLabelsState()
    {
      label_Filesreport.Click -= FilesReport_Click;
      label_Filesreport.Click += FilesReport_Click;
      label_Filesreport.Visible = true;
      label_Filesreport.Enabled = true;
      label_filesCode.Click -= addCode_Click;
      label_filesCode.Click += addCode_Click;
      label_filesCode.Visible = true;
      label_filesCode.Enabled = true;
    }

    private void CheckErrorNumeric()
    {
      if (numUD_NumberLab.Enabled && numUD_NumberLab.Value >= 1)
      {
        ActivateLabelsState();
      }
      else
      {
        ResetLabelsState();
      }
    }

    private void SetNumericUpDownEnabled(bool isEnabled)
    {
      numUD_NumberLab.Enabled = isEnabled;
      CheckNumberLabs();
    }

    private void numUD_NumberLab_ValueChanged(object sender, EventArgs e)
    {
      CheckNumberLabs();
    }
  }

  public class Group
  {
    public string Name { get; set; }
    public List<Student> Students { get; set; } = new List<Student>();
  }

  public class Student
  {
    public string Name { get; set; }
    public List<ReportData> Reports { get; set; } = new List<ReportData>();
  }

  public class ReportData
  {
    public string FileName { get; set; }
    public string FilePath { get; set; }
    public int WordCount { get; set; }
    public int SymbolCount { get; set; }
    public long AsciiSum { get; set; }
    public CodeAnalysis CodeInfo { get; set; } = null;
  }

  public class CodeAnalysis
  {
    public int CountFor { get; set; }
    public int CountWhile { get; set; }
    public int CountIf { get; set; }
    public int CountElse { get; set; }
  }
}
