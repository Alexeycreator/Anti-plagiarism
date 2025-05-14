using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace antiplagiat_lab
{
  public class LabFile
  {
    public string Id { get; set; }
    public string StudentName { get; set; }
    public string DocxFileName { get; set; }
    public string DocxFilePath { get; set; }
    public string TxtFilePath { get; set; }
    public long ASCII_Code { get; set; }
  }
}
