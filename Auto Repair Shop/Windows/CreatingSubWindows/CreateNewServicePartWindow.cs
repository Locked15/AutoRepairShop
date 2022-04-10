using System.Linq;
using System.Windows;
using Auto_Repair_Shop.Entities;

namespace Auto_Repair_Shop.Windows.CreatingSubWindows {

    /// <summary>
    /// Создает новую деталь.
    /// </summary>
    public partial class CreateNewServicePartWindow : Window {

        /// <summary>
        /// Новая запчасть.
        /// </summary>
        public Part newPart { get; set; }

        /// <summary>
        /// Конструктор класса.
        /// </summary>
        public CreateNewServicePartWindow() {
            InitializeComponent();
            newPart = new Part();

            DataContext = newPart;
        }

        #region Функции кнопок и прочие функции.

        /// <summary>
        /// Сбрасывает значения полей до значений по умолчанию.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void resetValues_Click(object sender, RoutedEventArgs e) {
            newPartName.Text = string.Empty;
            newPartPrice.Text = string.Empty;

            newPart = new Part();
            DataContext = newPart;
        }

        /// <summary>
        /// Создает новую деталь, если поля заполнены без ошибок.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void savePart_Click(object sender, RoutedEventArgs e) {
            if (checkToCorrect()) {
                DialogResult = true;

                Close();
            }
        }

        /// <summary>
        /// Проверяет корректность введенных данных.
        /// </summary>
        /// <returns>Корректность введенных данных.</returns>
        private bool checkToCorrect() {
            string error = string.Empty;

            if (string.IsNullOrEmpty(newPartName.Text))
                error += "Запчасти необходимо задать название.\n";
            else if (DBEntities.Instance.Parts.Any(x => x.Part_Name == newPart.Part_Name))
                error += "Запчасть с таким названием уже определена в системе.\n";

            if (string.IsNullOrEmpty(newPartPrice.Text))
                error += "Запчасти необходимо задать стоимость.\n";
            else if (newPart.Part_Price < 0)
                error += "Цена запчасти не может быть отрицательной.\n";

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
