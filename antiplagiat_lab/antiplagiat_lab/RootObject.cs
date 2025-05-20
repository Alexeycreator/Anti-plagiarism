using System.Collections.Generic;

namespace antiplagiat_lab
{
  public class RootObject
  {
    public List<string> Exclude { get; set; } = new List<string>();
    public Dictionary<string, GroupData> Groups { get; set; } = new Dictionary<string, GroupData>();
  }
}
