

using System.Collections.Generic;

namespace SharedProject.Models
{
    public class TestRailSection
    {
        public long id { get; set; }
        public string name { get; set; }
        public long depth { get; set; }
        public string description { get; set; }
        public long display_order { get; set; }
        public long? parent_id { get; set; }
        public long suite_id { get; set; }

    }
}
