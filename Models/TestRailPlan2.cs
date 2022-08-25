

using System.Collections.Generic;

namespace SharedProject.Models
{
    public class TestRailPlan2
    {
        public long suite_id { get; set; }
        public long assignedto_id { get; set; }
        public bool include_all { get; set; }
        public List<long> config_ids { get; set; } = new List<long>();
        public List<TestRailRun2> runs { get; set; } = new List<TestRailRun2>();

    }
}
