using System.Windows;
using Auto_Repair_Shop.Classes;

namespace Auto_Repair_Shop.Windows {
    
    /// <summary>
    /// Окно настроек.
    /// </summary>
    public partial class SettingsWindow : Window {
        /// <summary>
        /// Ссылка на текущие настройки приложения.
        /// </summary>
        public ProgramSettings settings = ProgramSettings.settings;

        /// <summary>
        /// Конструктор класса.
        /// </summary>
        public SettingsWindow() {
            InitializeComponent();
            DataContext = settings;
        }

        /// <summary>
        /// Сохраняет выбранные настройки в файл конфигурации.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы данного события.</param>
        private void saveSettings_Click(object sender, RoutedEventArgs e) {
            ProgramSettings.saveConfig();

            Close();
        }
    }
}
