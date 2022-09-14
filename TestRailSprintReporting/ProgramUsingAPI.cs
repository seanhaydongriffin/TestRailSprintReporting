using Newtonsoft.Json.Linq;
using SharedProject;
using SharedProject.Confluence;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Xml;
using System.Xml.Linq;

namespace TestRailSprintReporting
{
    class Program
    {
        private static SharedProject.TestRail.APIClient TestRailClient = null;

        static void Main(string[] args)
        {
            Log.Initialise(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\TestRailSprintReporting.log");
            Log.Initialise(null);
            AppConfig.Open();
            
            TestRailClient = new SharedProject.TestRail.APIClient(AppConfig.Get("TestRailUrl"));
            TestRailClient.User = AppConfig.Get("TestRailUser");
            TestRailClient.Password = AppConfig.Get("TestRailPassword");

            var today = System.DateTime.Today;
            var unixTimestamp = (int)today.Subtract(new System.DateTime(1970, 1, 1)).TotalSeconds;
            Log.WriteLine("today = " + today);
            Log.WriteLine("unixTimestamp = " + unixTimestamp);

            // test statuses

            //var test_statuses = (JArray)TestRailClient.SendGet("get_statuses");

            SharedProject.Confluence.APIClient ConfluenceClient = new SharedProject.Confluence.APIClient(AppConfig.Get("ConfluenceUrl"));
            ConfluenceClient.User = AppConfig.Get("ConfluenceUser");
            ConfluenceClient.Password = AppConfig.Get("ConfluenceApiToken");

//            var debug = (JObject)ConfluenceClient.SendGet(AppConfig.Get("ConfluencePageKey") + "?expand=body.storage");
            //var debug = (JObject)ConfluenceClient.SendGet("2196242885?expand=body.storage");
            //var debug = (JObject)ConfluenceClient.SendGet("2188705921?expand=body.storage");
            //var debug = (JObject)ConfluenceClient.SendGet("997687643/child/page?limit=200");
            //Log.WriteLine(debug.ToString());

            var current_quarter = SharedProject.DateTime.GetNowQuarterInfo();
            var current_sprint_name = "";
            System.DateTime current_sprint_start = new System.DateTime();
            System.DateTime current_sprint_end = new System.DateTime();
            var confluence_page_table = new Dictionary<string, DataTable>();
            var testrail_plan_type_status_count = new Dictionary<string, int?>();

            var daily_untested_count = new Dictionary<string, int?>();
            var daily_failed_count = new Dictionary<string, int?>();

            // The projects

            var project_num = 0;

            foreach (XmlNode TestProject in AppConfig.GetSectionGroup("TestProjects").GetSectionGroups())
            {
                project_num++;
                var TestProjectName = TestProject.GetAttributeValue("Name");
                var TestProjectId = TestProject.GetAttributeValue("Id");
                var TestProjectConfluenceRootKey = TestProject.GetAttributeValue("ConfluenceSpace") + "-" + TestProject.GetAttributeValue("ConfluencePage") + "-" + TestProject.GetAttributeValue("Team");

                if (!confluence_page_table.ContainsKey(TestProjectConfluenceRootKey))
                {
                    var results_table = new DataTable();
                    results_table.Columns.Add("TestProject", typeof(string));
                    results_table.Columns.Add("TestPlan", typeof(string));
                    results_table.Columns.Add("TestRun", typeof(string));
                    results_table.Columns.Add("TestAutoID", typeof(int));
                    results_table.Columns.Add("TestAutoName", typeof(string));
                    results_table.Columns.Add("TestTitle", typeof(string));
                    results_table.Columns.Add("TestStatus", typeof(string));
                    results_table.Columns.Add("TestTestedOn", typeof(string));
                    results_table.Columns.Add("TestAllDefects", typeof(string));
                    confluence_page_table.Add(TestProjectConfluenceRootKey, results_table);
                }

                // The quarter milestone

                Log.WriteLine("Prj " + project_num + " of " + AppConfig.GetSectionGroup("TestProjects").GetSectionGroups().Count + " \"" + TestProjectName + "\" getting the milestones ...");
                var milestones = (JObject)TestRailClient.SendGet("get_milestones/" + TestProjectId);
                var quarter_milestone = milestones.SelectToken("$..[?(@.name =~ /^FY" + current_quarter.ShortYear + "Q" + current_quarter.Quarter + " .*$/)]");

                if (quarter_milestone == null)

                    Environment.Exit(0);

                // the sprint milestone

                var sprint_milestones = quarter_milestone["milestones"];
                var sprint_milestone = sprint_milestones.SelectToken("$..[?(@.start_on <= " + unixTimestamp + " && @.due_on >= " + unixTimestamp + ")]");
                        
                if (sprint_milestone == null)

                    Environment.Exit(0);

                current_sprint_name = sprint_milestone["name"].ToString().Split(' ').FirstOrDefault();
                current_sprint_start = SharedProject.DateTime.UnixTimeStampToUTCDateTime(Convert.ToDouble(sprint_milestone["start_on"]));
                current_sprint_end = SharedProject.DateTime.UnixTimeStampToUTCDateTime(Convert.ToDouble(sprint_milestone["due_on"]));

                // The plans

                Log.WriteLine("Prj " + project_num + " of " + AppConfig.GetSectionGroup("TestProjects").GetSectionGroups().Count + " Milestone \"" + sprint_milestone["name"] + "\" getting the plans ...");

                var plans = (JObject)TestRailClient.SendGet("get_plans/" + TestProjectId + "&milestone_id=" + sprint_milestone["id"]);

                foreach (var plan in plans["plans"])
                {

                    // The runs

                    Log.WriteLine("Prj " + project_num + " of " + AppConfig.GetSectionGroup("TestProjects").GetSectionGroups().Count + " Plan \"" + plan["name"] + "\" getting the details ...");
                    plan["name"] = plan["name"].ToString().Replace(current_sprint_name, "").Trim();

                    var plan_detail = (JObject)TestRailClient.SendGet("get_plan/" + plan["id"]);

                    foreach (var entry in plan_detail["entries"])
                    {
                        foreach (var run in entry["runs"])
                        {
                            // The tests & results

                            Log.WriteLine("Prj " + project_num + " of " + AppConfig.GetSectionGroup("TestProjects").GetSectionGroups().Count + " Run \"" + run["name"] + " (" + run["config"] + ")\" getting the tests ...");

                            var tests = (JObject)TestRailClient.SendGet("get_tests/" + run["id"]);

                            Log.WriteLine("Prj " + project_num + " of " + AppConfig.GetSectionGroup("TestProjects").GetSectionGroups().Count + " Run \"" + run["name"] + " (" + run["config"] + ")\" getting the results ...");

                            var results = (JObject)TestRailClient.SendGet("get_results_for_run/" + run["id"]);
                            var tested_on = "";
                            var all_defects = "";

                            foreach (var test in tests["tests"])
                            {
                                try
                                {
                                    var test_results = results.SelectTokens("$.results[?(@.test_id == " + (long)test["id"] + ")]");
                                    var test_date_outcome = new Dictionary<string, string>();

                                    foreach (var test_result in test_results)
                                    {
//                                        var result_created_on = SharedProject.DateTime.UnixTimeStampToDateTime(Convert.ToDouble(test_result["created_on"])).ToString("ddd, dd/MM");
                                        var result_created_on = SharedProject.DateTime.UnixTimeStampToDateTime(Convert.ToDouble(test_result["created_on"])).ToString("ddd, dd/MM");

                                        if (!test_date_outcome.ContainsKey(result_created_on))

                                            test_date_outcome.put(result_created_on, AppConfig.Get("TestRailTestStatus" + test_result["status_id"]));
                                    }

                                    // Examine every day of the sprint for the test and increment the daily totals accordingly

                                    var prev_day_result = "Untested";

                                    for (var day = current_sprint_start; day.Date <= current_sprint_end; day = day.AddDays(1))
                                    {
                                        var day_str = day.ToString("ddd, dd/MM");

                                        if (!test_date_outcome.ContainsKey(day_str))

                                            test_date_outcome.put(day_str, prev_day_result);

                                        if (test_date_outcome.get(day_str).Equals("Untested"))
                                        
                                            daily_untested_count.Increment(plan["name"] + "~" + day_str);

                                        if (test_date_outcome.get(day_str).Equals("Failed"))
                                            
                                            daily_failed_count.Increment(plan["name"] + "~" + day_str);

                                        prev_day_result = test_date_outcome.get(day_str);
                                    }


                                    var latest_test_result = test_results.First();
                                    var created_on = SharedProject.DateTime.UnixTimeStampToDateTime((double)latest_test_result["created_on"]);
                                    tested_on = created_on.ToString("dd MMM HH:mm");
                                    all_defects = latest_test_result["defects"].ToString();
                                }
                                catch (Exception e)
                                {
                                }

                                //confluence_page_storage.Add(TestProjectConfluenceRootKey, "<tr><td><sub>" + test["custom_auto_script_ref"] + " (<a href=\"" + AppConfig.Get("TestRailUrl") + "/index.php?/cases/view/" + test["case_id"] + "\">C" + test["case_id"] + "</a>)</sub><ac:structured-macro ac:name=\"anchor\"><ac:parameter ac:name=\"\">C" + test["case_id"] + "</ac:parameter></ac:structured-macro></td><td><sub>" + Regex.Replace(Regex.Replace(test["title"].ToString().UrlEncode(5), "VARIANT on (\\w+)", "VARIANT on <a href=\"#$1\">$1</a>"), "DEPENDANT on (\\w+)", "DEPENDANT on <a href=\"#$1\">$1</a>") + "</sub></td><td><sub>" + status_emoticon + "<a href=\"" + AppConfig.Get("TestRailUrl") + "/index.php?/tests/view/" + test["id"] + "\">" + AppConfig.Get("TestRailTestStatus" + test["status_id"]) + "</a></sub></td><td><sub>" + tested_on + "</sub></td><td><sub>" + all_defects + "</sub></td></tr>");

//                                confluence_page_table[TestProjectConfluenceRootKey].Rows.Add(TestProjectName.ToString().Replace(" Assessment", "").Trim(), plan["name"].ToString().Replace(current_sprint_name, "").Trim(), run["name"].ToString().Replace(" Functional", "").Replace(" Regression", "").Trim(), test["case_id"], test["custom_auto_script_ref"], test["title"], AppConfig.Get("TestRailTestStatus" + test["status_id"]), tested_on, all_defects);
                                confluence_page_table[TestProjectConfluenceRootKey].Rows.Add(TestProjectName.ToString().Replace(" Assessment", "").Trim(), plan["name"], run["name"].ToString().Replace(" Functional", "").Replace(" Regression", "").Trim(), test["case_id"], test["custom_auto_script_ref"], test["title"], AppConfig.Get("TestRailTestStatus" + test["status_id"]), tested_on, all_defects);




                                // Tally the test results for later pie charting

                                testrail_plan_type_status_count.Increment(plan["name"].ToString().Replace(current_sprint_name, "").Trim() + "-" + AppConfig.Get("TestRailTestStatus" + test["status_id"]));
                            }
                        }
                    }
                }
            }

            // for each confluence page to update

            foreach (var each_confluence_page_table in confluence_page_table)
            {
                var confluence_space_key = each_confluence_page_table.Key.Split('-')[0];
                var confluence_parent_page_key = each_confluence_page_table.Key.Split('-')[1];
                var team_name = each_confluence_page_table.Key.Split('-')[2];

                // Check if the sprint page exists in Confluence (under the provided root page provided in the config)

                var confluence_child_page = (JObject)ConfluenceClient.SendGet(confluence_parent_page_key + "/child/page?limit=200");

                var sprint_page_key = confluence_child_page.SelectToken("$..results[?(@.title == '" + team_name + " " + current_sprint_name + "')].id");

                if (sprint_page_key == null)
                {
                    // Create the Sprint Page

                    Log.WriteLine("Confluence Page \"" + team_name + " " + current_sprint_name + "\" creating ...");

                    var confluence_create_page_json = new
                    {
                        type = "page",
                        title = team_name + " " + current_sprint_name,
                        @space = new
                        {
                            key = confluence_space_key
                        },
                        @ancestors = new[] {
                            new {
                                id = confluence_parent_page_key
                            }
                        }.ToList()
                    };

                    var result = (JObject)ConfluenceClient.SendPost("", confluence_create_page_json);
                    sprint_page_key = result["id"];
                }

                confluence_child_page = (JObject)ConfluenceClient.SendGet(sprint_page_key + "/child/page?limit=200");

                var sprint_qa_page_key = confluence_child_page.SelectToken("$..results[?(@.title == '" + team_name + " " + current_sprint_name + " QA Test Results')].id");

                if (sprint_qa_page_key == null)
                {
                    // Create the Sprint QA Page

                    Log.WriteLine("Confluence Page \"" + team_name + " " + current_sprint_name + " QA Test Results\" creating ...");

                    var confluence_create_page_json = new
                    {
                        type = "page",
                        title = team_name + " " + current_sprint_name + " QA Test Results",
                        @space = new
                        {
                            key = confluence_space_key
                        },
                        @ancestors = new[] {
                            new {
                                id = sprint_page_key
                            }
                        }.ToList()
                    };

                    var result = (JObject)ConfluenceClient.SendPost("", confluence_create_page_json);
                    sprint_qa_page_key = result["id"];
                }

                DataView results_table_view = new DataView(each_confluence_page_table.Value);
                
                // Table default sorting

                results_table_view.Sort = "TestPlan, TestProject, TestRun";
                var latest_regression_test_plan = "";
                var table_rows = new TableRows();

                foreach (DataRowView row in results_table_view)
                {
                    var emoji_name = "";
                    var emoji_shortname = "";
                    var emoji_id = "";
                    var emoji_fallback = "";

                    if (row["TestStatus"].Equals("Passed"))

                        emoji_name = "tick";

                    if (row["TestStatus"].Equals("Failed"))

                        emoji_name = "cross";

                    if (row["TestStatus"].Equals("Untested"))
                    {
                        emoji_name = "flag_off";
                        emoji_shortname = ":flag_off:";
                        emoji_id = "atlassian-flag_off";
                        emoji_fallback = ":flag_off:";
                    }

                    var status = new Emoticon(emoji_name, emoji_shortname, emoji_id, emoji_fallback);

                    table_rows.Add(
                        row["TestProject"],
                        row["TestPlan"],
                        row["TestRun"],
                        "[C" + row["TestAutoID"] + "] " + row["TestAutoName"],
                        row["TestTitle"],
                        status.GetXElement(),
                        row["TestTestedOn"],
                        row["TestAllDefects"]
                    );

                    if (row["TestPlan"].ToString().StartsWith("Regression"))

                        latest_regression_test_plan = row["TestPlan"].ToString();
                }

                var table = new Table("full-width").Add(
                    new TableColumnGroup(70, 130, 130, 140, 280, 40, 70, 90)).Add(
                    new TableHead("Client / Proj", "Test Type", "Test Level - Run", "[ID] Test Name", "Test Title", "Status", "Tested On", "All Defects")).Add(
                    new TableBody(table_rows));

                var confluence_page_storage_str = table.ToStringDisableFormatting();



                // Prepend burndown chart

                var burndown_chart_table_rows = new TableRows().AddHeader("Date", "Untested", "Failed");
                int? max_not_passed = 0;

                for (var day = current_sprint_start; day.Date <= current_sprint_end; day = day.AddDays(1))
                {
                    var day_str = day.ToString("ddd, dd/MM");

                    if (daily_untested_count.get(latest_regression_test_plan + "~" + day_str) == null)

                        daily_untested_count.put(latest_regression_test_plan + "~" + day_str, 0);

                    if (daily_failed_count.get(latest_regression_test_plan + "~" + day_str) == null)

                        daily_failed_count.put(latest_regression_test_plan + "~" + day_str, 0);

                    if (day.Date > System.DateTime.Today)
                    {
                        day_str = "(TBC) " + day_str;
                        daily_untested_count.put(latest_regression_test_plan + "~" + day_str, 0);
                        daily_failed_count.put(latest_regression_test_plan + "~" + day_str, 0);
                    }

                    if (max_not_passed < (daily_untested_count.get(latest_regression_test_plan + "~" + day_str) + daily_failed_count.get(latest_regression_test_plan + "~" + day_str)))

                        max_not_passed = daily_untested_count.get(latest_regression_test_plan + "~" + day_str) + daily_failed_count.get(latest_regression_test_plan + "~" + day_str);

                    burndown_chart_table_rows.Add(
                        day_str,
                        daily_untested_count.get(latest_regression_test_plan + "~" + day_str).ToString(),
                        daily_failed_count.get(latest_regression_test_plan + "~" + day_str).ToString()
                    );
                }

                var burndown_chart = new Chart(
                    new ChartBody(
                        new TableBody(burndown_chart_table_rows)
                    ),
                    latest_regression_test_plan + " Burndown Chart",
                    "bar",
                    "true",
                    500,
                    300,
                    "true",
                    "",
                    "up45",
                    "vertical",
                    Color.NameToHex("lightgray") + "," + Color.NameToHex("red"),
                    "",
                    max_not_passed
                );

                confluence_page_storage_str = burndown_chart.ToStringDisableFormatting() + confluence_page_storage_str;


                // Prepend pie charts (to top of) confluence page

                var section_width = "default";      // back to center
                //var section_width = "wide";      // go wide
                //var section_width = "full-width"; // go full wide
                var testrail_plan_type_status_processed = new Dictionary<string, bool?>();
                var pie_chart_columns = new SectionColumns();

                foreach (var entry in testrail_plan_type_status_count)
                {
                    var key = entry.Key.Substring(0, entry.Key.LastIndexOf('-'));

                    if (testrail_plan_type_status_processed.get(key) == null)
                    {
                        testrail_plan_type_status_processed.put(key, true);

                        pie_chart_columns.Add("center", new Chart(
                            new ChartBody(
                                new TableBody(
                                    new TableRows().AddHeader("Total", "Passed", "Failed", "Untested").Add(
                                        "Total",
                                        testrail_plan_type_status_count.get(key + "-Passed", 0).ToString(),
                                        testrail_plan_type_status_count.get(key + "-Failed", 0).ToString(),
                                        testrail_plan_type_status_count.get(key + "-Untested", 0).ToString()
                                    )
                                )
                            ),
                            key,
                            "pie",
                            "",
                            200,
                            200,
                            "false",
                            "",
                            "",
                            "",
                            Color.NameToHex("seagreen") + "," + Color.NameToHex("red") + "," + Color.NameToHex("lightgray"),
                            "%1%"
                        ));
                    }
                }

                var pie_chart_section = new Section("center", section_width, "true").Add(pie_chart_columns);

                //                confluence_page_storage_str = "<h5 style=\"text-align: center;\">Status Charts by Test Type</h5>" + pie_chart_section.ToString(SaveOptions.DisableFormatting) + confluence_page_storage_str;
                confluence_page_storage_str = "<h5 style=\"text-align: center;\">Status Charts by Test Type</h5>" + pie_chart_section.ToStringDisableFormatting() + confluence_page_storage_str;



                // Update the confluence page

                Log.WriteLine("Confluence Page \"" + sprint_qa_page_key + "\" getting the version ...");
                var confluence_page = (JObject)ConfluenceClient.SendGet(sprint_qa_page_key + "?expand=version");
                var confluence_page_version = (long)confluence_page["version"]["number"];
                confluence_page_version++;

                var confluence_json = new
                {
                    @version = new
                    {
                        number = confluence_page_version
                    },
                    type = "page",
                    title = team_name + " " + current_sprint_name + " QA Test Results",
                    @space = new
                    {
                        key = confluence_space_key
                    },
                    @ancestors = new[] {
                        new {
                            id = sprint_page_key
                        }
                    }.ToList(),
                    @body = new
                    {
                        @storage = new
                        {
                            value = confluence_page_storage_str,
                            representation = "storage"
                        }
                    }
                };

                Log.WriteLine("Confluence Page \"" + sprint_qa_page_key + "\" updating ...");
                var pp = (JToken)ConfluenceClient.SendPut(sprint_qa_page_key.ToString(), confluence_json);
                int i = 0;

            }
        }
    }
}
