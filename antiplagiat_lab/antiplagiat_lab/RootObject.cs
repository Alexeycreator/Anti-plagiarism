using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace antiplagiat_lab
{
  public class RootObject
  {
    public List<string> Exclude { get; set; }
    public Dictionary<string, GroupData> Groups { get; set; }
  }
}
