using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRefresher_Standalone
{
    internal class DataClasses
    {

        internal string Appfolder = Path.GetDirectoryName(Application.ExecutablePath) ?? string.Empty;



        internal string Config => Appfolder + @"\config.ini";
        internal string XML => Appfolder + @"\filepath.xml";
        


       

        internal string FetchLogPath()
        {


            string[] lines = File.ReadAllLines(Config);

            if (lines.Length > 0)
            {
                foreach (string line in lines)
                {
                    //MessageBox.Show(line);
                    if (line.StartsWith("-logFolder="))
                    {
                        string value = line.Split('=')[1].Trim();
                        return value;
                    }
                    else
                    {
                        return @"'-logFolder' Not Found ";
                    }

                }

            }
            else
            {
                return "Lines < 0";
            }
            return "Config Not Found";
        }
    }

}

