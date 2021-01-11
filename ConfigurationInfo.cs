using System;
using System.Collections.Generic;
using System.Linq;

namespace evaluacion_ceapsi
{
    class ConfigurationInfo
    {

        public const string CONFIG_FILE_DIR = @".\configs";
        public const string CONFIG_FILE_NAME = @".\configs\db";

        private static string[] existingConfigNames = { };


        public string Name { get; set; }
        public string TargetFile { get; set; }
        public string MappingFile { get; set; }

        public static void createNewConfiguration(string name, string targetFile, string mappingFile)
        {
            if (!System.IO.Directory.Exists(CONFIG_FILE_DIR))
            {
                System.IO.Directory.CreateDirectory(CONFIG_FILE_DIR);
            }
            string[] names = getExistingConfigNames();
            if (names.Contains(name))
            {
                throw new ArgumentException("Config name already exists", name);
            }
            bool append = System.IO.File.Exists(CONFIG_FILE_NAME);
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(CONFIG_FILE_NAME, append))
            {
                file.Write(name);
                file.WriteLine();
            }
            System.IO.File.Copy(targetFile, CONFIG_FILE_DIR + "\\" + name + ".xlsm");
            System.IO.File.Copy(mappingFile, CONFIG_FILE_DIR + "\\" + name + ".csv");
        }

        public static List<ConfigurationInfo> listConfigurations()
        {
            List<ConfigurationInfo> list = new List<ConfigurationInfo>();
            loadExistingNames();
            foreach (var name in existingConfigNames)
            {
                ConfigurationInfo info = new ConfigurationInfo
                {
                    Name = name,
                    TargetFile = String.Format("{0}\\{1}.xlsm", CONFIG_FILE_DIR, name),
                    MappingFile = String.Format("{0}\\{1}.csv", CONFIG_FILE_DIR, name)
                };
                list.Add(info);
            }
            return list;
        }

        public static string[] getExistingConfigNames()
        {
            if (existingConfigNames.Length == 0)
            {
                loadExistingNames();
            }
            return existingConfigNames;
        }

        private static void loadExistingNames()
        {
            if (System.IO.File.Exists(CONFIG_FILE_NAME))
            {
                System.IO.StreamReader sr = new System.IO.StreamReader(CONFIG_FILE_NAME);
                List<string> names = new List<string>();
                string name = sr.ReadLine();
                while (name != null)
                {
                    if (!name.Trim().Equals(String.Empty))
                    {
                        names.Add(name);
                    }
                    name = sr.ReadLine();
                }
                existingConfigNames = names.ToArray();
                sr.Close();
            }
            
        }

        public override string ToString()
        {
            return this.Name;
        }
    }
}
