

using System.Collections.Generic;

namespace SharedProject.Models
{
    public class TestRailPlan
    {
        public long id { get; set; }
        public string name { get; set; }
        public long suite_id { get; set; }
        public List<long> config_ids { get; set; } = new List<long>();
        public List<TestRailRun> runs { get; set; } = new List<TestRailRun>();

    }
}
