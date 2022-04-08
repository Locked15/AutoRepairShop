using System.IO;
using System.Text.Json;
using Auto_Repair_Shop.Resources;

namespace Auto_Repair_Shop.Classes { 

    /// <summary>
    /// Класс, содержащий настройки программы.
    /// </summary>
    public class ProgramSettings {

        #region Свойства класса.
        public static ProgramSettings settings { get; set; }

        public bool showCompletedRequests { get; set; }
        #endregion

        #region Конструкторы класса.
        /// <summary>
        /// Конструктор класса.
        /// </summary>
        public ProgramSettings() {
            showCompletedRequests = false;
        }

        /// <summary>
        /// Статический конструктор класса.
        /// </summary>
        static ProgramSettings() {
            var path = Path.Combine(ResourceManager.getCurrentPath(), "Config.json");

            if (File.Exists(path)) {
                using (StreamReader sr = new StreamReader(path, System.Text.Encoding.Default)) {
                    string content = sr.ReadToEnd();

                    settings = JsonSerializer.Deserialize<ProgramSettings>(content);
                }
            }

            else {
                settings = new ProgramSettings();
            }
        }
        #endregion

        #region Методы.
        /// <summary>
        /// Сохраняет текущую конфигурацию приложения.
        /// </summary>
        public static void saveConfig() {
            var path = Path.Combine(ResourceManager.getCurrentPath(), "Config.json");

            using (StreamWriter sw1 = new StreamWriter(path, false, System.Text.Encoding.Default)) {
                sw1.Write(JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true }));
            }
        }
        #endregion
    }
}
