using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn
{
    public static class Config
    {
        public static string UserDataFolder
        {
            get
            {
                string homeDir = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                string userDataFolder = Path.Combine(homeDir, "ExcelAddIn/UDF");
                Directory.CreateDirectory(userDataFolder);
                return userDataFolder;
            }
        }
    }
}
