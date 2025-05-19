using System;
using System.Collections.Generic;
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
    }

    private void InfoVariableForm_Load(object sender, EventArgs e)
    {
      Print();
    }

    private void Print()
    {
      rTbxInfoVariable.Text += $"{filePath}";
      rTbxInfoVariable.Text += $"\n{checkFileCode}";
      //try
      //{
      //  string code = File.ReadAllText(filePath);
      //  List<VariableInfo> variables = FindVariableDeclarations(code);

      //  rTbxInfoVariable.Text = "Найденные переменные:\n";
      //  foreach (var variable in variables)
      //  {
      //    rTbxInfoVariable.Text += variable.ToString() + "\n";
      //  }
      //}
      //catch (Exception ex)
      //{
      //  Console.WriteLine($"Ошибка: {ex.Message}");
      //}
    }

    public static List<VariableInfo> FindVariableDeclarations(string code)
    {
      var variables = new List<VariableInfo>();
      var declaredVariables = new HashSet<string>(); // Для отслеживания уже найденных переменных

      var patterns = new[]
{
    // Одиночные объявления с присваиванием (int a = 5;)
    new Regex(@"(?<type>\b[\w\.]+\b)\s+(?<name>\b\w+\b)\s*=\s*(?<value>[^;]+)\s*;"),

    // Var с присваиванием (var x = "text";)
    new Regex(@"\bvar\s+(?<name>\b\w+\b)\s*=\s*(?<value>[^;]+)\s*;"),

    // Множественные объявления (int a, b = 5, c;)
    new Regex(@"(?<type>\b[\w\.]+\b)\s+(?<names>(?:\s*\b\w+\b\s*(?:=\s*[^,;]+)?\s*,\s*)*\s*\b\w+\b\s*(?:=\s*[^,;]+)?)\s*;")
};

      foreach (var pattern in patterns)
      {
        var matches = pattern.Matches(code);
        foreach (Match match in matches)
        {
          if (match.Groups["names"].Success)
          {
            ProcessMultipleDeclarations(match, code, variables, declaredVariables);
          }
          else
          {
            ProcessSingleDeclaration(match, code, variables, declaredVariables);
          }
        }
      }
      return variables;
    }

    private static void ProcessSingleDeclaration(Match match, string code, List<VariableInfo> variables, HashSet<string> declaredVariables)
    {
      string varName = match.Groups["name"].Value;

      if (declaredVariables.Contains(varName))
        return;

      declaredVariables.Add(varName);

      // Если нет группы "value", то Value = null (для случаев типа "int a;")
      string varValue = match.Groups["value"]?.Success == true
          ? match.Groups["value"].Value
          : null;

      variables.Add(new VariableInfo
      {
        Type = match.Groups["type"]?.Value ?? "var",
        Name = varName,
        Value = varValue, // Будет null, если присваивания нет
        LineNumber = GetLineNumber(code, match.Index)
      });
    }

    private static void ProcessMultipleDeclarations(Match match, string code, List<VariableInfo> variables, HashSet<string> declaredVariables)
    {
      string type = match.Groups["type"].Value;
      string namesPart = match.Groups["names"].Value;

      // Разбиваем по запятым, убираем пробелы, игнорируем пустые
      var variableParts = namesPart.Split(',')
                                  .Select(part => part.Trim())
                                  .Where(part => !string.IsNullOrEmpty(part));

      foreach (var part in variableParts)
      {
        // Разделяем имя и значение (если есть "=")
        var nameValue = part.Split(new[] { '=' }, 2, StringSplitOptions.RemoveEmptyEntries);
        string varName = nameValue[0].Trim();
        string varValue = nameValue.Length > 1 ? nameValue[1].Trim() : null;

        if (declaredVariables.Contains(varName))
          continue;

        declaredVariables.Add(varName);

        variables.Add(new VariableInfo
        {
          Type = type,
          Name = varName,
          Value = varValue,
          LineNumber = GetLineNumber(code, match.Index)
        });
      }
    }


    private static int GetLineNumber(string code, int position)
    {
      if (position < 0 || position >= code.Length)
        return -1;

      int lineNumber = 1;
      for (int i = 0; i < position; i++)
      {
        if (code[i] == '\n')
          lineNumber++;
      }
      return lineNumber;
    }
  }
}
