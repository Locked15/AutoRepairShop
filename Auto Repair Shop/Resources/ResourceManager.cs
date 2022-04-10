using System;
using System.IO;

namespace Auto_Repair_Shop.Resources { 

    /// <summary>
    /// Работает с ресурсами проекта (изображениями, иконками и прочим).
    /// </summary>
    public class ResourceManager {

        /// <summary>
        /// Позволяет получить путь до директории проекта (папка, содержимое которой отображается в "Обозревателе решений").
        /// </summary>
        /// <returns>Путь к данной директории.</returns>
        public static string getCurrentPath() {
            string path = Environment.CurrentDirectory;

            for (int i = 0; i < 2; i++)
                path = path.Substring(0, path.LastIndexOf('\\'));

            return path;
        }

        /// <summary>
        /// Возвращает путь к изображению машины по умолчанию.
        /// </summary>
        /// <returns>Полный путь до изображения.</returns>
        public static string getDefaultImagePath() {
            string defaultPath = Path.Combine(getCurrentPath(), "Resources", "Pictures", "Default.png");

            return defaultPath;
        }

        /// <summary>
        /// Проверяет существование изображения с указанным путем.
        /// </summary>
        /// <param name="image">Путь к изображению.</param>
        /// <returns>Если изображение существует, вернется оригинальный путь. Если нет — путь к изображению по умолчанию.</returns>
        public static string checkExistsAndReturnFullPath(string image) {
            image = Path.Combine(getCurrentPath(), "Resources", "Pictures", image); 
            
            if (File.Exists(image)) {
                return image;
            } else {
                return getDefaultImagePath();
            }
        }

        /// <summary>
        /// Копирует существующее изображение с указанного пути в директорию ресурсов.
        /// </summary>
        /// <param name="fullPath">Полный путь до изображения, которое нужно добавить.</param>
        /// <returns>Относительный путь до изображения в ресурсах.</returns>
        public static string addStandaloneImageToResources(string fullPath) {
            string resourcesPath = Path.Combine(getCurrentPath(), "Resources", "Pictures", "Vehicles");
            string newImagePath = Path.Combine(resourcesPath, Path.GetFileName(fullPath));

            File.Copy(fullPath, newImagePath, true);

            return Path.Combine("Vehicles", Path.GetFileName(newImagePath));
        }
    }
}
