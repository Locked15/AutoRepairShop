using System;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using Auto_Repair_Shop.Entities;
using Auto_Repair_Shop.Resources;

namespace Auto_Repair_Shop.Windows.CreatingSubWindows {
    
    /// <summary>
    /// Окно добавления новой машины.
    /// </summary>
    public partial class AddNewVehicleWindow : Window {

        /// <summary>
        /// Новая машина.
        /// </summary>
        public Vehicle newVehicle { get; set; }

        #region Функции инициализации.

        /// <summary>
        /// Конструктор класса.
        /// </summary>
        public AddNewVehicleWindow() {
            newVehicle = new Vehicle();

            InitializeComponent();
            DataContext = newVehicle;

            initializeBrands();
            initializeClasses();
            initializeOwners();
            initializeImage();
        }

        /// <summary>
        /// Инициализирует поле с марками машин.
        /// </summary>
        private void initializeBrands() => vehicleBrand.ItemsSource = DBEntities.Instance.Vehicle_Brand.ToList();

        /// <summary>
        /// Инициализирует поле с классами машины.
        /// </summary>
        private void initializeClasses() => vehicleClass.ItemsSource = Enumerable.Range(1, 5).ToList();

        /// <summary>
        /// Инициализирует поле с возможными владельцами машины.
        /// </summary>
        private void initializeOwners() => vehicleOwner.ItemsSource = DBEntities.Instance.People.ToList();

        /// <summary>
        /// Инициализирует поле с изображением, устанавливая изображение по умолчанию.
        /// </summary>
        private void initializeImage() => vehicleImage.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri(ResourceManager.getDefaultImagePath()));
        #endregion

        #region Функции смены полей.

        /// <summary>
        /// Происходит при смене выбранного бренда автомобиля.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void vehicleBrand_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) {
            newVehicle.Vehicle_Brand = vehicleBrand.SelectedItem as Vehicle_Brand;
        }

        /// <summary>
        /// Происходит при смене класса машины.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void vehicleClass_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) {
            newVehicle.Vehicle_Class = Convert.ToInt32(vehicleClass.SelectedItem);
        }

        /// <summary>
        /// Происходит при смене выбранного владельца машины.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void vehicleOwner_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) {
            newVehicle.Person = vehicleOwner.SelectedItem as Person;
        }

        /// <summary>
        /// Открывает диалоговое окно установки нового изображения.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void setNewImage_Click(object sender, RoutedEventArgs e) {
            OpenFileDialog dialog = new OpenFileDialog()
            {
                Title = "Выбор изображения",
                Filter = "Изображения *.png|*.png|Изображения *.jpg|*.jpg|Изображения *.jpeg|*.jpeg|Все файлы|*.*",
                CheckFileExists = true,
                CheckPathExists = true,
                Multiselect = false
            };
            var result = dialog.ShowDialog();

            if (result == true) {
                newVehicle.Image = dialog.FileName;

                vehicleImage.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri(dialog.FileName));
            }
        }
        #endregion

        #region Функции сохранения машины в системе.

        /// <summary>
        /// Добавляет созданную машину в БД.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void saveVehicle_Click(object sender, RoutedEventArgs e) {
            if (checkToCorrect()) {
                newVehicle.Image = ResourceManager.addStandaloneImageToResources(newVehicle.Image);

                DialogResult = true;
                Close();
            }
        }

        /// <summary>
        /// Проверяет введенные поля на корректность.
        /// </summary>
        /// <returns>Корректны ли введенные данные.</returns>
        private bool checkToCorrect() {
            var error = string.Empty;

            if (newVehicle.Vehicle_Brand != null && !newVehicle.Name.Contains(newVehicle.Vehicle_Brand.Brand) && 
                newVehicle.Vehicle_Brand.Brand != "Прочие")
                error += "Название машины не соответствует бренду.\n";

            if (newVehicle.Person == null)
                error += "Не указан владелец машины.\n";

            if (newVehicle.State_Number == null || newVehicle.State_Number.Length != 9)
                error += "Введенный номер некорректен.\n";

            if (newVehicle.Vehicle_Class < 1)
                error += "Некорректный класс машины.\n";

            if (newVehicle.Vehicle_Brand == null)
                error += "Не указан бренд машины.\n";

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
