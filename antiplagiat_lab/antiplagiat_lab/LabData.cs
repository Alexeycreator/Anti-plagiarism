using System.Collections.Generic;

namespace antiplagiat_lab
{
  public class LabData
  {
    public string TitleLab { get; set; }
    public int NumberLab { get; set; }
    public List<LabFile> Files { get; set; } = new List<LabFile>();
  }
}
