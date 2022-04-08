using System;

namespace Auto_Repair_Shop.Resources { 

    public class ResourceManager {

        public static string getCurrentPath() {
            string path = Environment.CurrentDirectory;

            for (int i = 0; i < 2; i++)
                path = path.Substring(0, path.LastIndexOf('\\'));

            return path;
        }
    }
}
