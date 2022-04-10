using System.Windows;
using Auto_Repair_Shop.Entities;

namespace Auto_Repair_Shop.Windows.CreatingSubWindows {

    /// <summary>
    /// Окно добавления нового клиента.
    /// </summary>
    public partial class AddNewPersonWindow : Window {

        /// <summary>
        /// Новый клиент автомастерской.
        /// </summary>
        public Person newPerson { get; set; }

        /// <summary>
        /// Конструктор класса.
        /// </summary>
        public AddNewPersonWindow() {
            InitializeComponent();
            
            newPerson = new Person();
            DataContext = newPerson;
        }

        #region Функции добавления пользователя.

        /// <summary>
        /// Проверяет корректность данных и внедряет нового клиента в систему.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void createPerson_Click(object sender, RoutedEventArgs e) {
            if (checkToCorrect()) {
                DialogResult = true;

                Close();
            }
        }

        /// <summary>
        /// Проверяет корректность введенных данных.
        /// </summary>
        /// <returns>Логическое значение корректности данных.</returns>
        private bool checkToCorrect() {
            string error = string.Empty;

            if (string.IsNullOrEmpty(newPerson.Name))
                error += "Имя клиента не введено.\n";

             if (string.IsNullOrEmpty(newPerson.Last_Name)) 
                error += "Фамилия клиента не введена.\n";
            
            if (error == string.Empty) {
                return true;
            } else {
                MessageBox.Show($"Обнаружены ошибки:\n{error}", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);

                return false;
            }
        }
        #endregion
    }
}
