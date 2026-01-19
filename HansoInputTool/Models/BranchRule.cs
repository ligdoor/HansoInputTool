using System;

namespace HansoInputTool.Models
{
    public class BranchRule
    {
        public string Pattern { get; set; } = "";
        public string BranchGroup { get; set; } = "branch";
        public string NumberGroup { get; set; } = "number";
        public int Priority { get; set; } = 100;
    }
}
