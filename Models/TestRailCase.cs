

using System.Collections.Generic;

namespace SharedProject.Models
{
    public class TestRailCase
    {
        public long id { get; set; }
        public long section_id { get; set; }
        public string title { get; set; }
        public string full_title { get; set; }
        public string refs { get; set; }
        public bool? is_run_recently { get; set; }
        public List<TestRailCaseSteps> steps_separated { get; set; } = new List<TestRailCaseSteps>();

    }
}
