using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Auto_Repair_Shop.Entities;
using Auto_Repair_Shop.Windows.CreatingSubWindows;

namespace Auto_Repair_Shop.Windows.CreatingWindows {

    /// <summary>
    /// Окно добавления заказа.
    /// </summary>
    public partial class RequestSettingWindow : Window {

        #region Свойства.

        /// <summary>
        /// Заказ создается или изменяется?
        /// </summary>
        public bool creating { get; set; }

        /// <summary>
        /// Если здесь "true", после закрытия окна заказ будет удален.
        /// </summary>
        public bool toDelete { get; set; }

        /// <summary>
        /// Создаваемый/изменяемый заказ.
        /// </summary>
        public Service_Request request { get; set; }
        #endregion

        #region Функции инициализации.

        /// <summary>
        /// Конструктор класса.
        /// </summary>
        /// <param name="request">Заказ для изменения/создания.</param>
        public RequestSettingWindow(Service_Request request) {
            toDelete = false;
            creating = request == null;
            this.request = request ?? new Service_Request();

            InitializeComponent();
            DataContext = this;

            initializeVehiclesList();
            initializeRequesterList();
            initializeFinalDate();
            initializeServiceTypeList();

            Title = creating ? "Создание заказа — Auto Repair Shop" : $"{request.Vehicle.Name} — Auto Repair Shop";
        }

        /// <summary>
        /// Инициализирует поле выбора автомобиля.
        /// </summary>
        private void initializeVehiclesList() {
            var allVehicles = DBEntities.Instance.Vehicles.ToList();

            vehicleSelect.ItemsSource = allVehicles;
            vehicleSelect.SelectedIndex = allVehicles.FindIndex(x => x.Id == request.Vehicle_Id);
        }

        /// <summary>
        /// Инициализирует поле выбора заказчика.
        /// </summary>
        private void initializeRequesterList() {
            var allPersons = DBEntities.Instance.People.ToList();

            requesterSelect.ItemsSource = allPersons;
            requesterSelect.SelectedIndex = allPersons.FindIndex(x => x.Id == (request.Person == null ? -1 : request.Person.Id));
        }

        /// <summary>
        /// Инициализирует поле выбора конечной даты, устанавливая минимальное значение.
        /// </summary>
        private void initializeFinalDate() => completeDateSelect.DisplayDateStart = requestDateSelect.SelectedDate;

        /// <summary>
        /// Инициализирует поле выбора типа услуги.
        /// </summary>
        private void initializeServiceTypeList() {
            var serviceTypes = DBEntities.Instance.Service_Type.ToList();

            selectServiceType.ItemsSource = serviceTypes;
            selectServiceType.SelectedIndex = serviceTypes.FindIndex(x => x.Id == request.Service_Type_Id);
        }
        #endregion

        #region События смены значений полей.

        /// <summary>
        /// Происходит при смене выбранного автомобиля. 
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void vehicleSelect_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            request.Vehicle = vehicleSelect.SelectedItem as Vehicle;

            Title = $"{request.Vehicle.Name} — Auto Repair Shop";
        }

        /// <summary>
        /// Происходит при смене заказчика.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void requesterSelect_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            request.Person = requesterSelect.SelectedItem as Person;
        }

        /// <summary>
        /// Происходит при смене даты размещения заказа.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void requestDateSelect_SelectedDateChanged(object sender, SelectionChangedEventArgs e) {
            request.Request_Date = requestDateSelect.SelectedDate;
            completeDateSelect.DisplayDateStart = requestDateSelect.SelectedDate;

            if (request.Request_Approx_Complete < request.Request_Date) {
                request.Request_Approx_Complete = request.Request_Date;

                completeDateSelect.SelectedDate = request.Request_Approx_Complete;
            }
        }

        /// <summary>
        /// Происходит при смене даты выполнения заказа.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void completeDateSelect_SelectedDateChanged(object sender, SelectionChangedEventArgs e) => request.Request_Approx_Complete = completeDateSelect.SelectedDate;

        /// <summary>
        /// Происходит при смене выбранного типа услуги.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void selectServiceType_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            request.Service_Type = selectServiceType.SelectedItem as Service_Type;

            // Проверка на типы услуг, которые требуют запчасти для выполнения.
            if (request.Service_Type.Id == 1 || request.Service_Type.Id == 4 || request.Service_Type.Id == 5) {
                setRequestParts.Visibility = Visibility.Visible;
            } else {
                setRequestParts.Visibility = Visibility.Hidden;
            }
        }
        #endregion

        #region События нажатия кнопок.

        /// <summary>
        /// Выбирает текущий день в поле выбора дня размещения заказа.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void selectTodayInRequestDate_Click(object sender, RoutedEventArgs e) => requestDateSelect.SelectedDate = DateTime.Now;

        /// <summary>
        /// Выбирает текущий день в поле выбора дня выполнения заказа.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void selectTodayInCompleteDate_Click(object sender, RoutedEventArgs e) {
            if (request.Request_Date < DateTime.Now) {
                completeDateSelect.SelectedDate = DateTime.Now;
            } else {
                MessageBox.Show("Несоответствие дат: заказ выполнен раньше, чем размещён.", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Открывает окно выбора деталей для текущего заказа.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void setRequestParts_Click(object sender, RoutedEventArgs e) {

        }
        #endregion

        #region События контрольной панели.

        /// <summary>
        /// Открывает окно создания автомобиля.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void createVehicle_Click(object sender, RoutedEventArgs e) {
            AddNewVehicleWindow window = new AddNewVehicleWindow();
            var result = window.ShowDialog();

            if (result.HasValue && result.Value) {
                DBEntities.Instance.Vehicles.Add(window.newVehicle);
                DBEntities.Instance.SaveChanges();

                initializeVehiclesList();

                vehicleSelect.SelectedIndex = DBEntities.Instance.Vehicles.ToList().FindIndex(v => v.Id == request.Vehicle_Id);
            }
        }

        /// <summary>
        /// Открывает окно добавления человека в систему.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события</param>
        private void createPerson_Click(object sender, RoutedEventArgs e) {
            AddNewPersonWindow window = new AddNewPersonWindow();
            var result = window.ShowDialog();

            if (result.HasValue && result.Value) {
                DBEntities.Instance.People.Add(window.newPerson);
                DBEntities.Instance.SaveChanges();

                initializeRequesterList();

                requesterSelect.SelectedIndex = DBEntities.Instance.People.ToList().FindIndex(v => v.Id == request.Requester_Id);
            }
        }

        /// <summary>
        /// Помечает заказ на удаление. 
        /// <br/>
        /// Запрашивает подтверждение пользователя и удаляет заказ при подтверждении.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void deleteRequest_Click(object sender, RoutedEventArgs e) {

            if (!creating) {
                var result = MessageBox.Show("Вы точно хотите удалить заказ?\n\nЭто действие необратимо.", "Внимание", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes) {
                    DialogResult = true;
                    toDelete = true;

                    Close();
                }
            } else {
                MessageBox.Show("Невозможно удалить ещё не созданный заказ.", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Сохраняет заказ в системе.
        /// <br/>
        /// Если он создается, то он будет добавлен. В ином случае просто будут сохранены изменения.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события</param>
        private void saveRequest_Click(object sender, RoutedEventArgs e) {
            if (checkCorrect()) {
                if (request.Service_Type.Id != 1 && request.Service_Type.Id != 4 && request.Service_Type.Id != 5) {
                    request.Parts_To_Request.Clear();
                }

                DialogResult = true;
                Close();
            }
        }
        #endregion

        #region Прочие функции.

        /// <summary>
        /// Проверяет корректность текущего состояния заказа.
        /// </summary>
        /// <returns>Корректность заказа.</returns>
        private bool checkCorrect() {
            string error = string.Empty;

            if (!request.Request_Date.HasValue)
                error += "Необходимо добавить дату размещения заказа.\n";

            if (request.Request_Approx_Complete < request.Request_Date)
                error += "Дата выполнения заказа не может быть раньше его размещения.\n";

            if (request.Vehicle == null)
                error += "Не выбран автомобиль для обслуживания.\n";

            if (request.Person == null)
                error += "Не указан заказчик.\n";

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
