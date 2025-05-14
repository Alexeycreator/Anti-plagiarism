namespace antiplagiat_lab
{
  public class VariableInfo
  {
    public string Type { get; set; }
    public string Name { get; set; }
    public string Value { get; set; }
    public int LineNumber { get; set; }

    public override string ToString()
    {
      return Value != null
          ? $"Строка {LineNumber}: {Type} {Name} = {Value}"
          : $"Строка {LineNumber}: {Type} {Name}";
    }
  }
}
