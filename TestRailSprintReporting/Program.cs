using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SharedProject;
using SharedProject.Models;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using Toolkit.Windows;

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

            // var debug = (JObject)ConfluenceClient.SendGet(AppConfig.Get("ConfluencePageKey") + "?expand=body.storage");
            //var debug = (JObject)ConfluenceClient.SendGet("2188705921?expand=body.storage");
            //var debug = (JObject)ConfluenceClient.SendGet("997687643/child/page?limit=200");
            //Log.WriteLine(debug.ToString());

            var current_quarter = SharedProject.DateTime.GetNowQuarterInfo();
            var current_sprint = "";
            NameValueCollection confluence_page_storage = new NameValueCollection();
            var testrail_plan_type_status_count = new Dictionary<string, int?>();

            // The projects

            var project_num = 0;

            foreach (XmlNode TestProject in AppConfig.GetSectionGroup("TestProjects").GetSectionGroups())
            {
                project_num++;
                var TestProjectName = TestProject.GetAttributeValue("Name");
                var TestProjectId = TestProject.GetAttributeValue("Id");
                var TestProjectConfluenceRootKey = TestProject.GetAttributeValue("ConfluenceSpace") + "-" + TestProject.GetAttributeValue("ConfluencePage") + "-" + TestProject.GetAttributeValue("Team");

                if (confluence_page_storage[TestProjectConfluenceRootKey] == null)
                {
                    //confluence_page_storage.Add(TestProjectConfluenceRootKey, "<ac:structured-macro ac:name=\"info\"><ac:rich-text-body>Current Sprint is <strong>" + current_sprint_with_date + "</strong></ac:rich-text-body></ac:structured-macro>");
                    confluence_page_storage.Add(TestProjectConfluenceRootKey, "<ac:structured-macro ac:name=\"toc\" ac:schema-version=\"1\" data-layout=\"default\"><ac:parameter ac:name=\"minLevel\">1</ac:parameter><ac:parameter ac:name=\"maxLevel\">2</ac:parameter></ac:structured-macro>");
                    confluence_page_storage.Add(TestProjectConfluenceRootKey, "<hr />");
                }

                confluence_page_storage.Add(TestProjectConfluenceRootKey, "<h1><u><a href=\"" + AppConfig.Get("TestRailUrl") + "/index.php?/projects/overview/" + TestProjectId + "\">" + TestProjectName + "</a></u></h1>");

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

                current_sprint = sprint_milestone["name"].ToString().Split(' ').FirstOrDefault();

                // The plans

                Log.WriteLine("Prj " + project_num + " of " + AppConfig.GetSectionGroup("TestProjects").GetSectionGroups().Count + " Milestone \"" + sprint_milestone["name"] + "\" getting the plans ...");

                var plans = (JObject)TestRailClient.SendGet("get_plans/" + TestProjectId + "&milestone_id=" + sprint_milestone["id"]);

                foreach (var plan in plans["plans"])
                {
                    confluence_page_storage.Add(TestProjectConfluenceRootKey, "<h2><u><a href=\"" + AppConfig.Get("TestRailUrl") + "/index.php?/plans/view/" + plan["id"] + "\">" + plan["name"].ToString().Replace(current_sprint, "").Trim() + "</a></u></h2>");

                    // The runs

                    Log.WriteLine("Prj " + project_num + " of " + AppConfig.GetSectionGroup("TestProjects").GetSectionGroups().Count + " Plan \"" + plan["name"] + "\" getting the details ...");

                    var plan_detail = (JObject)TestRailClient.SendGet("get_plan/" + plan["id"]);

                    foreach (var entry in plan_detail["entries"])
                    {
                        foreach (var run in entry["runs"])
                        {
                            confluence_page_storage.Add(TestProjectConfluenceRootKey, "<h3><a href=\"" + AppConfig.Get("TestRailUrl") + "/index.php?/runs/view/" + run["id"] + "\">" + run["name"] + " (" + run["config"] + ")</a></h3>");
                            confluence_page_storage.Add(TestProjectConfluenceRootKey, "<table data-layout=\"full-width\"><colgroup><col style=\"width:110px;\"/><col style=\"width:280px;\"/><col style=\"width:50px;\"/><col style=\"width:60px;\"/><col style=\"width:100px;\"/></colgroup><tbody><tr><td><sub><b>Name (ID)</b></sub></td><td><sub><b>Title</b></sub></td><td><sub><b>Status</b></sub></td><td><sub><b>Tested On</b></sub></td><td><sub><b>All Defects</b></sub></td></tr>");

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
                                    var latest_test_result = results.SelectTokens("$.results[?(@.test_id == " + (long)test["id"] + ")]").First();
                                    var created_on = SharedProject.DateTime.UnixTimeStampToDateTime((double)latest_test_result["created_on"]);
                                    tested_on = created_on.ToString("dd MMM h:mm tt");
                                    all_defects = latest_test_result["defects"].ToString();
                                }
                                catch (Exception e)
                                {
                                }

                                var status_emoticon = "";

                                if (AppConfig.Get("TestRailTestStatus" + test["status_id"]).Equals("Passed"))

                                    status_emoticon = "<ac:emoticon ac:name=\"tick\" /> ";

                                if (AppConfig.Get("TestRailTestStatus" + test["status_id"]).Equals("Failed"))

                                    status_emoticon = "<ac:emoticon ac:name=\"cross\" /> ";

                                if (AppConfig.Get("TestRailTestStatus" + test["status_id"]).Equals("Untested"))

                                    status_emoticon = "<ac:emoticon ac:name=\"flag_off\" ac:emoji-shortname=\":flag_off:\" ac:emoji-id=\"atlassian-flag_off\" ac:emoji-fallback=\":flag_off:\" /> ";

                                confluence_page_storage.Add(TestProjectConfluenceRootKey, "<tr><td><sub>" + test["custom_auto_script_ref"] + " (<a href=\"" + AppConfig.Get("TestRailUrl") + "/index.php?/cases/view/" + test["case_id"] + "\">C" + test["case_id"] + "</a>)</sub><ac:structured-macro ac:name=\"anchor\"><ac:parameter ac:name=\"\">C" + test["case_id"] + "</ac:parameter></ac:structured-macro></td><td><sub>" + Regex.Replace(Regex.Replace(test["title"].ToString().UrlEncode(5), "VARIANT on (\\w+)", "VARIANT on <a href=\"#$1\">$1</a>"), "DEPENDANT on (\\w+)", "DEPENDANT on <a href=\"#$1\">$1</a>") + "</sub></td><td><sub>" + status_emoticon + "<a href=\"" + AppConfig.Get("TestRailUrl") + "/index.php?/tests/view/" + test["id"] + "\">" + AppConfig.Get("TestRailTestStatus" + test["status_id"]) + "</a></sub></td><td><sub>" + tested_on + "</sub></td><td><sub>" + all_defects + "</sub></td></tr>");

                                // Tally the test results for later pie charting

                                testrail_plan_type_status_count.Increment(plan["name"].ToString().Replace(current_sprint, "").Trim() + "-" + AppConfig.Get("TestRailTestStatus" + test["status_id"]));
                            }

                            confluence_page_storage.Add(TestProjectConfluenceRootKey, "</tbody></table>");
                        }
                    }
                }
            }

            // for each confluence page to update

            foreach (string confluence_root_key in confluence_page_storage)
            {
                var confluence_space_key = confluence_root_key.Split('-')[0];
                var confluence_parent_page_key = confluence_root_key.Split('-')[1];
                var team_name = confluence_root_key.Split('-')[2];

                // Check if the sprint page exists in Confluence (under the provided root page provided in the config)

                var confluence_child_page = (JObject)ConfluenceClient.SendGet(confluence_parent_page_key + "/child/page?limit=200");

                var sprint_page_key = confluence_child_page.SelectToken("$..results[?(@.title == '" + team_name + " " + current_sprint + "')].id");

                if (sprint_page_key == null)
                {
                    // Create the Sprint Page

                    Log.WriteLine("Confluence Page \"" + team_name + " " + current_sprint + "\" creating ...");

                    var confluence_create_page_json = new
                    {
                        type = "page",
                        title = team_name + " " + current_sprint,
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

                var sprint_qa_page_key = confluence_child_page.SelectToken("$..results[?(@.title == '" + team_name + " " + current_sprint + " QA')].id");

                if (sprint_qa_page_key == null)
                {
                    // Create the Sprint QA Page

                    Log.WriteLine("Confluence Page \"" + team_name + " " + current_sprint + " QA\" creating ...");

                    var confluence_create_page_json = new
                    {
                        type = "page",
                        title = team_name + " " + current_sprint + " QA",
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

                var confluence_page_storage_str = string.Join("", confluence_page_storage.GetValues(confluence_root_key));

                // Prepend pie charts (to top of) confluence page

                var section_width = "default";      // back to center
                //var section_width = "wide";      // go wide
                //var section_width = "full-width"; // go full wide
                var chart_size = 200;
                var pie_charts_storage_str = "<ac:structured-macro ac:name=\"section\" ac:align=\"center\" data-layout=\"" + section_width + "\"><ac:parameter ac:name=\"border\">true</ac:parameter><ac:rich-text-body>";
                var testrail_plan_type_status_processed = new Dictionary<string, bool?>();

                foreach (var entry in testrail_plan_type_status_count)
                {
                    var key = entry.Key.Substring(0, entry.Key.LastIndexOf('-'));

                    if (testrail_plan_type_status_processed.get(key) == null)
                    {
                        testrail_plan_type_status_processed.put(key, true);

                        var pie_chart = "<ac:structured-macro ac:name=\"chart\"><ac:parameter ac:name=\"subTitle\">" + key + "</ac:parameter><ac:parameter ac:name=\"type\">pie</ac:parameter><ac:parameter ac:name=\"width\">" + chart_size + "</ac:parameter><ac:parameter ac:name=\"height\">" + chart_size + "</ac:parameter><ac:parameter ac:name=\"legend\">false</ac:parameter><ac:parameter ac:name=\"pieSectionLabel\">%1%</ac:parameter><ac:parameter ac:name=\"colors\">green,red,gray</ac:parameter><ac:rich-text-body><table><tbody>" +
                                        "<tr><th><p>Total</p></th><th><p>Passed</p></th><th><p>Failed</p></th><th><p>Untested</p></th></tr>" +
                                        "<tr><th><p>Total</p></th><td><p>" + testrail_plan_type_status_count.get(key + "-Passed", 0) + "</p></td><td><p>" + testrail_plan_type_status_count.get(key + "-Failed", 0) + "</p></td><td><p>" + testrail_plan_type_status_count.get(key + "-Untested", 0) + "</p></td></tr>" +
                                        "</tbody></table></ac:rich-text-body></ac:structured-macro>";

                        pie_charts_storage_str = pie_charts_storage_str + "<ac:structured-macro ac:name=\"column\" ac:align=\"center\"><ac:rich-text-body>" + pie_chart + "</ac:rich-text-body></ac:structured-macro>";
                    }
                }

                pie_charts_storage_str = pie_charts_storage_str + "</ac:rich-text-body></ac:structured-macro>";
                confluence_page_storage_str = pie_charts_storage_str + confluence_page_storage_str;

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
                    title = team_name + " " + current_sprint + " QA",
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

            }
        }
    }
}
