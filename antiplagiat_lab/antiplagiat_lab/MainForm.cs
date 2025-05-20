using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace antiplagiat_lab
{
  public partial class MainForm : Form
  {
    private const string DataFilePath = "data.json";
    private const string ReportsDirectory = "Отчёты";
    private List<Group> groups = new List<Group>();
    private int selectedIndexNumberLab;
    private int selectedIndexGroup;
    private int selectedIndexStudent;
    private static string _docxFilePath;
    private static string _txtFilePath;
    private static string _docxFileName;
    private static long _ascii_code;
    private ReportData reportData;
    private string typeFile;
    public string checkCoincidenceFileCode;
    public string checkSelectedStudentFileCode;

    public MainForm()
    {
      InitializeComponent();
      InitializeApp();
    }

    #region actions_Btn

    private void numUD_NumberLab_Enter(object sender, EventArgs e)
    {
      selectedIndexGroup = comboBox_Group.SelectedIndex;
      selectedIndexStudent = comboBox_Student.SelectedIndex;
      selectedIndexNumberLab = Convert.ToInt32(numUD_NumberLab.Value);
      if (!string.IsNullOrWhiteSpace(comboBox_currentReport.Text))
      {
        comboBox_Group.SelectedIndex = -1;
        comboBox_Student.SelectedIndex = -1;
        comboBox_currentReport.SelectedIndex = -1;
      }
    }

    private void numUD_NumberLab_ValueChanged(object sender, EventArgs e)
    { // отчистка при изменении в нумерик
      label_CountWords.Text = null;
      label_CountSymbol.Text = null;
      label_SummASCII.Text = null;
      dataGridView_Coincidence.Rows.Clear();
      CheckNumberLabs();
      if (!numUD_NumberLab.Enabled)
      {
        return;
      }
      if (string.IsNullOrEmpty(comboBox_currentReport.Text))
      {
        if (numUD_NumberLab.Value == selectedIndexNumberLab)
        {
          comboBox_Group.SelectedIndex = selectedIndexGroup;
          comboBox_Student.SelectedIndex = selectedIndexStudent;
        }
        else
        {
          comboBox_Group.SelectedIndex = -1;
          comboBox_Student.SelectedIndex = -1;
        }
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
      label_countFOR.Text = "       ";
      label_countWHILE.Text = "       ";
      label_countIF.Text = "       ";
      label_countELSE.Text = "       ";
      label_CountWords.Text = "       ";
      label_CountSymbol.Text = "       ";
      label_SummASCII.Text = "       ";
      dataGridView_Coincidence.Rows.Clear();
      UpdateStudentList();
      comboBox_Student.SelectedItem = null;
      comboBox_currentReport.SelectedItem = null;
      comboBox_Student.Enabled = true;
    }

    private void FilesReport_Click(object sender, EventArgs e)
    {
      try
      {
        typeFile = "doc";
        if (comboBox_Group.SelectedItem != null && comboBox_Student.SelectedItem != null)
        {
          using (OpenFileDialog openFileDialog = new OpenFileDialog())
          {
            openFileDialog.Multiselect = true; // Включаем множественный выбор файлов
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

                // Очищаем комбобокс перед добавлением новых отчетов
                comboBox_currentReport.Items.Clear();

                // Обрабатываем все выбранные файлы
                foreach (string filePath in openFileDialog.FileNames)
                {
                  string destFile = Path.Combine(Environment.CurrentDirectory, destPath, Path.GetFileName(filePath));
                  bool fileExists = student.Reports.Any(r => r.FileName == Path.GetFileName(filePath));
                  string fileName = Path.GetFileName(filePath);
                  string destFilePath = Path.Combine(destPath, fileName);

                  _docxFilePath = destFilePath;
                  _docxFileName = fileName;

                  if (fileExists)
                  {
                    var existingReport = student.Reports.First(r => r.FileName == Path.GetFileName(filePath));

                    var dialogResult = MessageBox.Show(
                        $"Отчет с таким названием уже существует.\n\n" +
                        $"Текущий файл: {existingReport.FileName}\n" +
                        $"Размер: {existingReport.SymbolCount} символов, {existingReport.WordCount} слов\n" +
                        $"Хотите заменить его новым файлом?\n\n" +
                        $"Новый файл: {Path.GetFileName(filePath)}",
                        "Заменить файл?",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question
                    );

                    if (dialogResult == DialogResult.Yes)
                    {
                      student.Reports.Remove(existingReport);
                      File.Copy(filePath, destFile, true);
                      var newReportData = AnalyzeReport(destFile);
                      student.Reports.Add(newReportData);
                      comboBox_currentReport.Items.Add(newReportData.FileName);
                    }
                  }
                  else
                  {
                    File.Copy(filePath, destFile, true);
                    var newReportData = AnalyzeReport(destFile);
                    student.Reports.Add(newReportData);
                    comboBox_currentReport.Items.Add(newReportData.FileName);
                  }
                }

                // Сохраняем данные после обработки всех файлов
                SaveData();

                // Выбираем последний добавленный отчет
                if (comboBox_currentReport.Items.Count > 0)
                {
                  comboBox_currentReport.SelectedIndex = comboBox_currentReport.Items.Count - 1;
                }
              }
            }
          }
        }
        else
        {
          MessageBox.Show("Пожалуйста, выберите группу и студента.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        SerializationDataLabs();
      }
      catch (Exception ex)
      {
        MessageBox.Show($"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
    }

    private void addCode_Click(object sender, EventArgs e)
    {
      typeFile = "text";
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
              _txtFilePath = Path.GetFullPath(openFileDialog.FileName);
              return;
            }
          }

          string code = File.ReadAllText(openFileDialog.FileName);
          report.CodeInfo = AnalyzeCode(code);
          string relativePath = Path.Combine("Отчёты", selectedGroup, selectedStudent, "Code.txt");
          _txtFilePath = relativePath;
          SerializationDataLabs();
          SaveData();
          DisplayCodeInfo(report.CodeInfo);
          try
          {
            string targetDir = Path.Combine("Отчёты", selectedGroup, selectedStudent);
            Directory.CreateDirectory(targetDir);
            string targetPath = Path.Combine(targetDir, "Code.txt");
            File.Copy(openFileDialog.FileName, targetPath, true);
          }
          catch (Exception ex)
          {
            MessageBox.Show($"Ошибка при копировании файла: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
          }
        }
      }
    }

    private void InformationVariableToolStripMenuItem_Click(object sender, EventArgs e)
    {
      InfoVariableForm infoVariableForm = new InfoVariableForm(checkSelectedStudentFileCode, checkCoincidenceFileCode);
      infoVariableForm.ShowDialog();
    }

    private void button_Open_Click(object sender, EventArgs e)
    {
      if (dataGridView_Coincidence.SelectedRows.Count > 0)
      {
        string filePath = dataGridView_Coincidence.SelectedRows[0].Cells[6].Value.ToString();
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
      label_countFOR.Text = "       ";
      label_countWHILE.Text = "       ";
      label_countIF.Text = "       ";
      label_countELSE.Text = "       ";
      label_CountWords.Text = "       ";
      label_CountSymbol.Text = "       ";
      label_SummASCII.Text = "       ";
      dataGridView_Coincidence.Rows.Clear();
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


    private void SetNumericUpDownEnabled(bool isEnabled)
    {
      numUD_NumberLab.Enabled = isEnabled;
      CheckNumberLabs();
    }

    #endregion

    #region Functions

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
        groups = System.Text.Json.JsonSerializer.Deserialize<List<Group>>(jsonData) ?? new List<Group>();
      }
    }

    private void SerializationDataLabs()
    {
      try
      {
        RootObject root;
        string filePath = "dataCheckLabs.json";

        if (File.Exists(filePath))
        {
          string existingJson = File.ReadAllText(filePath);
          root = JsonConvert.DeserializeObject<RootObject>(existingJson);
        }
        else
        {
          root = new RootObject();
          root.Groups = new Dictionary<string, GroupData>
          {
            ["ПМИ"] = new GroupData(),
            ["МКН"] = new GroupData(),
            ["АИС"] = new GroupData()
          };
        }
        int labNumber = Convert.ToInt32(numUD_NumberLab.Value);
        string labKey = $"Labs_{labNumber}";
        string groupName = comboBox_Group.Text;
        string studentName = comboBox_Student.Text;
        if (!root.Groups.ContainsKey(groupName))
        {
          root.Groups[groupName] = new GroupData();
        }

        if (!root.Groups[groupName].Fields.ContainsKey(labKey))
        {
          root.Groups[groupName].Fields[labKey] = new LabData
          {
            TitleLab = $"Лабораторная_работа_{labNumber}",
            NumberLab = labNumber,
            Files = new List<LabFile>()
          };
        }
        var currentLab = root.Groups[groupName].Fields[labKey];
        var existingStudentFiles = currentLab.Files.Where(f => f.StudentName == studentName).ToList();
        if (existingStudentFiles.Any())
        {
          switch (typeFile)
          {
            case "doc":
              var firstDialogResult = MessageBox.Show(
              $"Студент {studentName} уже существует в работе {labNumber}\n" +
              $"Добавление в базу данных проверок невозможно.",
              "Ошибка",
              MessageBoxButtons.OKCancel,
              MessageBoxIcon.Error);

              if (firstDialogResult == DialogResult.OK)
              {
                var secondDialogResult = MessageBox.Show(
                    $"Хотите пересоздать данные о студенте {studentName} в группе {groupName}?",
                    "Выберите действие",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (secondDialogResult == DialogResult.Yes)
                {
                  foreach (var file in existingStudentFiles)
                  {
                    currentLab.Files.Remove(file);
                  }
                  string _jsonString = JsonConvert.SerializeObject(root, Formatting.Indented);
                  File.WriteAllText(filePath, _jsonString);
                }
                else
                {
                  return;
                }
              }
              else
              {
                return;
              }
              break;
            case "text":
              var _firstDialogResult = MessageBox.Show(
              $"Файл исходного кода у студента {studentName} уже существует в работе {labNumber}\n" +
              $"Добавление в базу данных проверок невозможно.",
              "Ошибка",
              MessageBoxButtons.OKCancel,
              MessageBoxIcon.Error);

              if (_firstDialogResult == DialogResult.OK)
              {
                var _secondDialogResult = MessageBox.Show(
                    $"Хотите пересоздать данные о файле исходного кода студента {studentName} в группе {groupName}?",
                    "Выберите действие",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (_secondDialogResult == DialogResult.Yes)
                {
                  foreach (var file in existingStudentFiles)
                  {
                    currentLab.Files.Remove(file);
                  }
                  string __jsonString = JsonConvert.SerializeObject(root, Formatting.Indented);
                  File.WriteAllText(filePath, __jsonString);
                }
                else
                {
                  return;
                }
              }
              else
              {
                return;
              }
              break;
            default: break;
          }
        }

        switch (typeFile)
        {
          case "doc":
            currentLab.Files.Add(new LabFile
            {
              Id = Guid.NewGuid().ToString(),
              StudentName = comboBox_Student.Text,
              DocxFileName = _docxFileName,
              DocxFilePath = _docxFilePath,
              TxtFilePath = null,
              ASCII_Code = _ascii_code
            });
            if (_docxFileName != "Otchet.docx" || _docxFileName != "Otchet.docx")
            {
              throw new FormatException($"Загруженный файл имеет некорректное название. Исправьте на 'Otchet'!\n" +
                $"Не исправив название, будет невозможно проверить исходный код данного студента.");
            }
            break;
          case "text":
            const string jsonFilePath = "dataCheckLabs.json";
            var jsonData = File.Exists(jsonFilePath)
                ? JsonConvert.DeserializeObject<RootObject>(File.ReadAllText(jsonFilePath))
                : new RootObject();
            string _groupName = comboBox_Group.Text;
            string labName = $"Labs_{labNumber}";
            string reportsFolder = "Отчёты";
            if (!jsonData.Groups.ContainsKey(groupName))
            {
              jsonData.Groups[groupName] = new GroupData { Fields = new Dictionary<string, LabData>() };
            }
            if (!jsonData.Groups[groupName].Fields.ContainsKey(labName))
            {
              jsonData.Groups[groupName].Fields[labName] = new LabData { Files = new List<LabFile>() };
            }
            var lab = jsonData.Groups[groupName].Fields[labName];
            var existingJsonFile = lab.Files.Find(f => string.Equals(f.StudentName, studentName, StringComparison.Ordinal));
            string docxFileName = null;
            string docxFilePath = null;

            if (existingJsonFile != null)
            {
              if (!string.IsNullOrEmpty(existingJsonFile.DocxFileName))
              {
                docxFileName = existingJsonFile.DocxFileName;
                if (!docxFileName.Equals("Otchet.docx", StringComparison.OrdinalIgnoreCase) &&
                    !docxFileName.Equals("Otchet.doc", StringComparison.OrdinalIgnoreCase))
                {
                  throw new FormatException("Название файла должно быть 'Otchet.docx' или 'Otchet.doc'");
                }
                docxFilePath = !string.IsNullOrEmpty(existingJsonFile.DocxFilePath)
                  ? existingJsonFile.DocxFilePath.Contains(reportsFolder)
                  ? existingJsonFile.DocxFilePath.Substring(existingJsonFile.DocxFilePath.IndexOf(reportsFolder))
                  : Path.Combine(reportsFolder, groupName, studentName, docxFileName ?? "Otchet.docx")
                  : Path.Combine(reportsFolder, groupName, studentName, docxFileName ?? "Otchet.docx");
              }
            }
            if (string.IsNullOrEmpty(docxFileName))
            {
              string docxPath = Path.Combine(reportsFolder, groupName, studentName, "Otchet.docx");
              string docPath = Path.Combine(reportsFolder, groupName, studentName, "Otchet.doc");

              if (File.Exists(docxPath))
              {
                docxFileName = "Otchet.docx";
                docxFilePath = docxPath;
              }
              else if (File.Exists(docPath))
              {
                docxFileName = "Otchet.doc";
                docxFilePath = docPath;
              }
              else
              {
                docxFileName = "Otchet.docx";
                docxFilePath = Path.Combine(reportsFolder, groupName, studentName, docxFileName);
              }
            }

            long asciiCode = 0;
            bool asciiCodeValid = false;

            if (!string.IsNullOrWhiteSpace(label_SummASCII.Text))
            {
              string cleanAscii = new string(label_SummASCII.Text.Where(c => char.IsDigit(c)).ToArray());

              if (!string.IsNullOrEmpty(cleanAscii) && long.TryParse(cleanAscii, out long parsedAscii))
              {
                asciiCode = parsedAscii;
                asciiCodeValid = true;
              }
            }
            if (!asciiCodeValid && existingJsonFile != null)
            {
              asciiCode = existingJsonFile.ASCII_Code;
            }
            var existingFile = currentLab.Files.Find(f => string.Equals(f.StudentName, studentName, StringComparison.Ordinal));

            if (existingFile != null)
            {
              existingFile.TxtFilePath = _txtFilePath;
              existingFile.DocxFileName = docxFileName;
              existingFile.DocxFilePath = docxFilePath;
              existingFile.ASCII_Code = asciiCode;
            }
            else
            {
              var newFile = new LabFile
              {
                Id = Guid.NewGuid().ToString(),
                StudentName = studentName,
                TxtFilePath = _txtFilePath,
                DocxFileName = docxFileName,
                DocxFilePath = docxFilePath,
                ASCII_Code = asciiCode
              };

              currentLab.Files.Add(newFile);
              lab.Files.Add(new LabFile
              {
                Id = newFile.Id,
                StudentName = newFile.StudentName,
                DocxFileName = newFile.DocxFileName,
                DocxFilePath = newFile.DocxFilePath,
                TxtFilePath = newFile.TxtFilePath,
                ASCII_Code = newFile.ASCII_Code
              });
            }
            File.WriteAllText(jsonFilePath, JsonConvert.SerializeObject(jsonData, Formatting.Indented));
            break;
          default: break;
        }
        string jsonString = JsonConvert.SerializeObject(root, Formatting.Indented);
        File.WriteAllText(filePath, jsonString);
        MessageBox.Show("Данные успешно сохранены!");
      }
      catch (FormatException ex)
      {
        MessageBox.Show($"{ex.Message}", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
      }
      catch (Exception ex)
      {
        MessageBox.Show($"Ошибка сохранения: {ex.Message}");
      }
    }

    private void SaveData()
    {
      string jsonData = System.Text.Json.JsonSerializer.Serialize(groups);
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
    private ReportData AnalyzeReport(string filePath)
    {
      try
      {
        if (!File.Exists(filePath))
        {
          File.Delete(filePath);
          throw new FileNotFoundException("Файл не найден.");
        }
        if (Path.GetExtension(filePath) == ".doc" || Path.GetExtension(filePath) == ".docx")
        {
          MessageBox.Show($"Формат файла выбран верно");
        }
        else if (Path.GetExtension(filePath) != ".docx")
        {
          File.Delete(filePath);
          throw new ArgumentException("Поддерживаются только форматы .doc или .docx");
        }
        else if (Path.GetExtension(filePath) != ".doc")
        {
          File.Delete(filePath);
          throw new ArgumentException("Поддерживаются только форматы .doc или .docx");
        }

        Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
        Document doc = null;
        reportData = null;

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
          _ascii_code = asciiSum;
          reportData = new ReportData
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
      catch (FileNotFoundException ex)
      {
        MessageBox.Show($"{ex.Message}", "Ошибка поиска файла", MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
      catch (ArgumentException ex)
      {
        MessageBox.Show($"{ex.Message}", "Ошибка формата файла", MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
      catch (Exception ex)
      {
        MessageBox.Show($"{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
      return reportData;
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
      string studentCodePath = Path.Combine("Отчёты", comboBox_Group.Text, comboBox_Student.Text, "Code.txt");
      if (codeInfo == null)
      {
        label_countFOR.Text = "       ";
        label_countWHILE.Text = "       ";
        label_countIF.Text = "       ";
        label_countELSE.Text = "       ";
      }
      else if (studentCodePath != null)
      {
        if (File.Exists(studentCodePath))
        {
          label_countFOR.Text = $"{codeInfo.CountFor}";
          label_countWHILE.Text = $"{codeInfo.CountWhile}";
          label_countIF.Text = $"{codeInfo.CountIf}";
          label_countELSE.Text = $"{codeInfo.CountElse}";
        }
      }
      else
      {
        label_countFOR.Text = "       ";
        label_countWHILE.Text = "       ";
        label_countIF.Text = "       ";
        label_countELSE.Text = "       ";
      }
    }

    private int CountOccurrences(string text, string pattern)
    {
      return System.Text.RegularExpressions.Regex.Matches(text, pattern).Count;
    }

    private void FillDataGridView(long currentAsciiSum, string currentFilePath)
    {
      dataGridView_Coincidence.Rows.Clear();
      checkCoincidenceFileCode = string.Empty;
      checkSelectedStudentFileCode = string.Empty;
      try
      {
        string jsonPath = "dataCheckLabs.json";
        if (!File.Exists(jsonPath))
        {
          MessageBox.Show($"Файл данных для проверки не найден");
          return;
        }
        string jsonData = File.ReadAllText(jsonPath);
        RootObject root = JsonConvert.DeserializeObject<RootObject>(jsonData);
        string selectGroup = comboBox_Group.Text;
        if (!root.Groups.ContainsKey(selectGroup))
        {
          MessageBox.Show($"Группа {selectGroup} не найдена в данных!");
          return;
        }
        int currentLabNumber = Convert.ToInt32(numUD_NumberLab.Value);
        string labKeyToCompare = $"Labs_{currentLabNumber}";

        if (!root.Groups[selectGroup].Fields.ContainsKey(labKeyToCompare))
        {
          MessageBox.Show($"Лабораторная работа №{currentLabNumber} не найдена в группе {selectGroup}!");
          return;
        }
        var filesToCompare = root.Groups[selectGroup].Fields[labKeyToCompare].Files;
        string currentStudent = comboBox_Student.Text;
        bool hasMatches = false;
        foreach (var file in filesToCompare)
        {
          if (file.StudentName == currentStudent)
          {
            checkSelectedStudentFileCode = file.TxtFilePath;
          }
          if (file.DocxFilePath == currentFilePath || file.StudentName == currentStudent)
          {
            continue;
          }
          double similarityPercentage = CalculateSimilarityPercentage(currentAsciiSum, file.ASCII_Code);
          if (similarityPercentage <= 25)
          {
            dataGridView_Coincidence.Rows.Add(
                file.StudentName,
                selectGroup,
                currentLabNumber,
                file.ASCII_Code,
                Math.Round(100 - similarityPercentage, 2),
                file.DocxFileName,
                file.DocxFilePath,
                file.TxtFilePath
            );
            hasMatches = true;
          }
        }
        if (hasMatches && dataGridView_Coincidence.Rows.Count > 0)
        {
          var txtFilePathCell = dataGridView_Coincidence.Rows[0].Cells[7];
          if (txtFilePathCell.Value != null)
          {
            checkCoincidenceFileCode = txtFilePathCell.Value.ToString();
          }
        }
      }
      catch (Exception ex)
      {
        MessageBox.Show($"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
    }

    private double CalculateSimilarityPercentage(long currentAsciiSum, long reportAsciiSum)
    {
      double maxAsciiSum = Math.Max(currentAsciiSum, reportAsciiSum);
      double percentage = (Math.Abs(currentAsciiSum - reportAsciiSum) / maxAsciiSum) * 100;

      return percentage;
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
      dataGridView_Coincidence.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
      dataGridView_Coincidence.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
      dataGridView_Coincidence.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
      this.AutoScroll = true;
      this.MouseWheel += new MouseEventHandler(MainForm_MouseWheel);
    }

    private void MainForm_MouseWheel(object sender, MouseEventArgs e)
    {
      if (e.Delta > 0)
      {
        verticalScrollBar.Value = Math.Max(verticalScrollBar.Value - verticalScrollBar.SmallChange * 3, verticalScrollBar.Minimum);
      }
      else
      {
        verticalScrollBar.Value = Math.Min(verticalScrollBar.Value + verticalScrollBar.SmallChange * 3, verticalScrollBar.Maximum);
      }
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

    #endregion

  }
  #region Class
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
  #endregion
}
