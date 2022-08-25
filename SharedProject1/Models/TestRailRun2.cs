

using System.Collections.Generic;

namespace SharedProject.Models
{
    public class TestRailRun2
    {
        public bool include_all { get; set; }
        public List<long> case_ids { get; set; } = new List<long>();
        public List<long> config_ids { get; set; } = new List<long>();

    }
}
