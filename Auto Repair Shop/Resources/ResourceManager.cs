using System;
using System.IO;

namespace Auto_Repair_Shop.Resources { 

    public class ResourceManager {

        public static string getCurrentPath() {
            string path = Environment.CurrentDirectory;

            for (int i = 0; i < 2; i++)
                path = path.Substring(0, path.LastIndexOf('\\'));

            return path;
        }

        /// <summary>
        /// Проверяет существование изображения с указанным путем.
        /// </summary>
        /// <param name="image">Путь к изображению.</param>
        /// <returns>Если изображение существует, вернется оригинальный путь. Если нет — путь к изображению по умолчанию.</returns>
        public static string checkExistsAndReturnFullPath(string image) {
            if (File.Exists(image)) {
                return image;
            } else {
                string defaultPath = Path.Combine(getCurrentPath(), "Resources", "Default.png");

                return defaultPath;
            }
        }
    }
}
