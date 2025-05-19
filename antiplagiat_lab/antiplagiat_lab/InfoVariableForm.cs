using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace antiplagiat_lab
{
    public partial class InfoVariableForm : Form
    {
        private readonly string filePath;
        private readonly string checkFileCode;

        public InfoVariableForm(string _filePath, string _checkFileCode)
        {
            filePath = _filePath;
            checkFileCode = _checkFileCode;
            InitializeComponent();
            SetupUI();
        }

        private void SetupUI()
        {
            rTbxInfoVariable.Dock = DockStyle.Fill;
            rTbxInfoVariable.Font = new Font("Consolas", 10);
            rTbxInfoVariable.ReadOnly = true;
            rTbxInfoVariable.ScrollBars = RichTextBoxScrollBars.Both;
            rTbxInfoVariable.WordWrap = false;

            var compareButton = new Button();
            compareButton.Text = "Сравнить переменные";
            compareButton.Dock = DockStyle.Top;
            compareButton.Height = 40;
            compareButton.Click += CompareButton_Click;
            this.Controls.Add(compareButton);
        }

        private void CompareButton_Click(object sender, EventArgs e)
        {
            CompareVariables();
        }

        private void CompareVariables()
        {
            try
            {
                rTbxInfoVariable.Clear();
                AppendColoredText("Начало сравнения переменных...\n", Color.DarkBlue);

                // Получаем переменные из первого файла (filePath)
                List<VariableInfo> firstVars = new List<VariableInfo>();
                if (File.Exists(filePath))
                {
                    firstVars = GetVariablesFromFile(filePath);
                    AppendColoredText($"\nФайл 1: {filePath}", Color.Black);
                    AppendColoredText($" - найдено {firstVars.Count} переменных\n", Color.DarkGreen);
                }
                else
                {
                    AppendColoredText($"\nФайл 1: {filePath} не найден\n", Color.Red);
                }

                // Получаем переменные из второго файла (checkFileCode)
                List<VariableInfo> secondVars = new List<VariableInfo>();
                if (File.Exists(checkFileCode))
                {
                    secondVars = GetVariablesFromFile(checkFileCode);
                    AppendColoredText($"Файл 2: {checkFileCode}", Color.Black);
                    AppendColoredText($" - найдено {secondVars.Count} переменных\n", Color.DarkGreen);
                }
                else
                {
                    AppendColoredText($"\nФайл 2: {checkFileCode} не найден\n", Color.Red);
                }

                // Сравниваем переменные
                var matches = FindExactMatches(firstVars, secondVars);
                DisplayMatches(matches, filePath, checkFileCode);

                AppendColoredText("\nСравнение завершено!\n", Color.DarkBlue);
            }
            catch (Exception ex)
            {
                AppendColoredText($"\nОШИБКА: {ex.Message}\n", Color.Red);
            }
        }

        private List<VariableInfo> GetVariablesFromFile(string path)
        {
            string code = File.ReadAllText(path);
            return ParseVariables(code);
        }

        private List<VariableInfo> ParseVariables(string code)
        {
            var variables = new List<VariableInfo>();
            var declaredVars = new HashSet<string>();

            var patterns = new[]
            {
                new Regex(@"(?<type>\b[\w\.]+\b)\s+(?<name>\b\w+\b)\s*=\s*(?<value>[^;]+)\s*;"),
                new Regex(@"\bvar\s+(?<name>\b\w+\b)\s*=\s*(?<value>[^;]+)\s*;"),
                new Regex(@"(?<type>\b[\w\.]+\b)\s+(?<names>(?:\s*\b\w+\b\s*(?:=\s*[^,;]+)?\s*,\s*)*\s*\b\w+\b\s*(?:=\s*[^,;]+)?)\s*;")
            };

            foreach (var pattern in patterns)
            {
                foreach (Match match in pattern.Matches(code))
                {
                    if (match.Groups["names"].Success)
                    {
                        ProcessMultiDeclaration(match, code, variables, declaredVars);
                    }
                    else
                    {
                        ProcessSingleDeclaration(match, code, variables, declaredVars);
                    }
                }
            }

            return variables;
        }

        private void ProcessSingleDeclaration(Match match, string code, List<VariableInfo> variables, HashSet<string> declaredVars)
        {
            string varName = match.Groups["name"].Value;
            if (declaredVars.Contains(varName)) return;

            declaredVars.Add(varName);
            variables.Add(new VariableInfo
            {
                Type = match.Groups["type"]?.Value ?? "var",
                Name = varName,
                Value = match.Groups["value"]?.Value,
                LineNumber = GetLineNumber(code, match.Index)
            });
        }

        private void ProcessMultiDeclaration(Match match, string code, List<VariableInfo> variables, HashSet<string> declaredVars)
        {
            string type = match.Groups["type"].Value;
            foreach (var part in match.Groups["names"].Value.Split(',').Select(p => p.Trim()).Where(p => !string.IsNullOrEmpty(p)))
            {
                var nameValue = part.Split(new[] { '=' }, 2);
                string varName = nameValue[0].Trim();
                if (declaredVars.Contains(varName)) continue;

                declaredVars.Add(varName);
                variables.Add(new VariableInfo
                {
                    Type = type,
                    Name = varName,
                    Value = nameValue.Length > 1 ? nameValue[1].Trim() : null,
                    LineNumber = GetLineNumber(code, match.Index)
                });
            }
        }

        private int GetLineNumber(string code, int pos)
        {
            return code.Substring(0, pos).Count(c => c == '\n') + 1;
        }

        private List<VariableMatch> FindExactMatches(List<VariableInfo> firstVars, List<VariableInfo> secondVars)
        {
            var matches = new List<VariableMatch>();
            var secondDict = secondVars.GroupBy(v => (v.Name, v.Type)).ToDictionary(g => g.Key, g => g.ToList());

            foreach (var firstVar in firstVars)
            {
                if (secondDict.TryGetValue((firstVar.Name, firstVar.Type), out var secondList))
                {
                    matches.AddRange(secondList.Select(secondVar => new VariableMatch
                    {
                        VariableName = firstVar.Name,
                        Type = firstVar.Type,
                        FirstFileLine = firstVar.LineNumber,
                        SecondFileLine = secondVar.LineNumber,
                        Value = firstVar.Value
                    }));
                }
            }

            return matches;
        }

        private void DisplayMatches(List<VariableMatch> matches, string file1, string file2)
        {
            AppendColoredText($"\nНайдено совпадений: {matches.Count}\n", matches.Count > 0 ? Color.DarkGreen : Color.DarkOrange);

            for (int i = 0; i < matches.Count; i++)
            {
                var m = matches[i];
                AppendColoredText($"\nСовпадение #{i + 1}:\n", Color.Purple);
                AppendColoredText($"  Переменная: ", Color.Black);
                AppendColoredText($"{m.VariableName}\n", Color.DarkCyan);
                AppendColoredText($"  Тип: ", Color.Black);
                AppendColoredText($"{m.Type}\n", Color.DarkBlue);
                AppendColoredText($"  Расположение: ", Color.Black);
                AppendColoredText($"{Path.GetFileName(file1)} (строка {m.FirstFileLine})", Color.Blue);
                AppendColoredText(" ↔ ", Color.Gray);
                AppendColoredText($"{Path.GetFileName(file2)} (строка {m.SecondFileLine})\n", Color.Blue);

                if (!string.IsNullOrEmpty(m.Value))
                {
                    AppendColoredText($"  Значение: ", Color.Black);
                    AppendColoredText($"{m.Value}\n", Color.DarkGreen);
                }

                AppendColoredText(new string('-', 60) + "\n", Color.LightGray);
            }
        }

        private void AppendColoredText(string text, Color color)
        {
            rTbxInfoVariable.SelectionStart = rTbxInfoVariable.TextLength;
            rTbxInfoVariable.SelectionLength = 0;
            rTbxInfoVariable.SelectionColor = color;
            rTbxInfoVariable.AppendText(text);
            rTbxInfoVariable.SelectionColor = rTbxInfoVariable.ForeColor;
        }
        private void InfoVariableForm_Load(object sender, EventArgs e)
        {

        }
        public class VariableInfo
        {
            public string Type { get; set; }
            public string Name { get; set; }
            public string Value { get; set; }
            public int LineNumber { get; set; }
        }

        public class VariableMatch
        {
            public string VariableName { get; set; }
            public string Type { get; set; }
            public int FirstFileLine { get; set; }
            public int SecondFileLine { get; set; }
            public string Value { get; set; }
        }
    }

}

