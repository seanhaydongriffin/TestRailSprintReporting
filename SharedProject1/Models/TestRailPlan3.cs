

using System.Collections.Generic;

namespace SharedProject.Models
{
    public class TestRailPlan3
    {
        public string name { get; set; }
        public string description { get; set; }
        public long milestone_id { get; set; }
        public List<TestRailPlanEntry> entries { get; set; } = new List<TestRailPlanEntry>();

    }
}
