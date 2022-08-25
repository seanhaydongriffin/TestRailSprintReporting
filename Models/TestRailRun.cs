

using System.Collections.Generic;

namespace SharedProject.Models
{
    public class TestRailRun
    {
        public long id { get; set; }
        public string name { get; set; }
        public bool include_all { get; set; }
        public List<long> case_ids { get; set; } = new List<long>();
        public List<long> config_ids { get; set; } = new List<long>();

    }
}
