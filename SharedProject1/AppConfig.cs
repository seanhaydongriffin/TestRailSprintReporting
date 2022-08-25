using System;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Xml;

namespace SharedProject
{
    public static class AppConfig
    {
        static public Configuration configuration;
        static public ConfigurationSectionGroup BatchesConfigurationSectionGroup;
        static private string ConfigPath;

//        static public void Open(string config_filename = "AzureVMControl.config")
        static public void Open()
        {
            //            ConfigPath = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\" + config_filename;
            ConfigPath = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\TestRailSprintReporting.config";

            //if (!File.Exists(CoPilotConfigPath))

            //    CoPilotConfigPath = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\..\\..\\..\\AzureVMControl\\App.config";

            if (!File.Exists(ConfigPath))
            {
                //MessageBox.Show("Can't find CoPilot config.  Exiting.");
                Log.WriteLine("");
                Log.WriteLine("A valid configuration filename must be provided.");
                Environment.Exit(0);
            }

            System.Configuration.ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
            fileMap.ExeConfigFilename = ConfigPath;
            configuration = System.Configuration.ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None);

            BatchesConfigurationSectionGroup = configuration.GetSectionGroup("Batches");


        }

        static public XmlNode GetSection(this ConfigurationSectionGroup SectionGroup)
        {
            try
            {
                return Xml.GetNode(ConfigPath, "configuration/" + SectionGroup.SectionGroupName);
            }
            catch (Exception)
            {
            }

            return null;
        }

        static public XmlNodeList GetSectionGroups(this ConfigurationSectionGroup SectionGroup)
        {
            try
            {
                return Xml.GetChildNodes(ConfigPath, "configuration/" + SectionGroup.SectionGroupName + "/*");
            }
            catch (Exception)
            {
            }

            return null;
        }

        static public ConfigurationSectionGroup GetSectionGroup(string SectionGroupName)
        {
            try
            {
                return configuration.GetSectionGroup(SectionGroupName);
            }
            catch (Exception)
            {
            }

            return null;
        }

        static public XmlNodeList GetSectionGroups(string SectionGroupXPath)
        {
            try
            {
                return Xml.GetChildNodes(ConfigPath, SectionGroupXPath);
            }
            catch (Exception)
            {
            }

            return null;
        }

        static public XmlNodeList GetSectionGroups(ConfigurationSection Section)
        {
            try
            {
                var sectionXmlDoc2 = new XmlDocument();
                var xml2 = Section.SectionInformation.GetRawXml();
                sectionXmlDoc2.Load(new StringReader(xml2));
                return sectionXmlDoc2.SelectNodes("/*/add");
            }
            catch (Exception)
            {
            }

            return null;
        }


        static public XmlNodeList GetSectionSettings(ConfigurationSection Section)
        {
            try
            {
                var sectionXmlDoc2 = new XmlDocument();
                var xml2 = Section.SectionInformation.GetRawXml();
                sectionXmlDoc2.Load(new StringReader(xml2));
                return sectionXmlDoc2.SelectNodes("/*/add");
                //XmlNode titleNode2 = sectionXmlDoc2.SelectSingleNode("/*/add[@Hostname='azureauto8']");
            }
            catch (Exception)
            {
            }

            return null;
        }

        static public XmlNode GetSectionSetting(ConfigurationSection Section, string XPath)
        {
            try
            {
                var sectionXmlDoc2 = new XmlDocument();
                var xml2 = Section.SectionInformation.GetRawXml();
                sectionXmlDoc2.Load(new StringReader(xml2));

                // XPath = "/*/add[@Hostname='azureauto8']"
                return sectionXmlDoc2.SelectSingleNode(XPath);
            }
            catch (Exception)
            {
            }

            return null;
        }


        static public string Get(string key)
        {
            try
            {
                return configuration.AppSettings.Settings[key].Value;
            }
            catch (Exception)
            {
            }

            return null;
        }

        static public void Put(string key, string value)
        {
            configuration.AppSettings.Settings[key].Value = value;
        }

        static public void Save()
        {
            configuration.Save();
        }
    }
}
