using System;
using System.Linq;
using System.Windows;
using System.Collections.Generic;
using Auto_Repair_Shop.Entities;
using Auto_Repair_Shop.Windows.CreatingWindows;

namespace Auto_Repair_Shop.Windows.CreatingSubWindows {

    /// <summary>
    /// Окно добавления запчастей в заказ.
    /// </summary>
    public partial class AddServicePartsToRequestWindow : Window {

        #region Свойства.

        /// <summary>
        /// Список с деталями для выполнения заказа.
        /// </summary>
        public List<Parts_To_Request> partsToRequest { get; set; }

        /// <summary>
        /// Родительское окно.
        /// <br/>
        /// Оно нужно для формирования экземпляров "Parts_To_Request".
        /// </summary>
        private RequestSettingWindow parentWindow { get; set; }

        /// <summary>
        /// Содержит все связи, которые удаляются в процессе редактирования.
        /// <br/>
        /// EF не умеет автоматически удалять "висящие" связи, поэтому это нужно сделать вручную, а иначе — исключения.
        /// </summary>
        private List<Parts_To_Request> partsToRemove { get; set; } = new List<Parts_To_Request> (1);
        #endregion

        #region Функции инициализации.

        /// <summary>
        /// Конструктор класса.
        /// </summary>
        /// <param name="parts">Список деталей заказа.</param>
        public AddServicePartsToRequestWindow(RequestSettingWindow parent, List<Parts_To_Request> parts) {
            InitializeComponent();

            DataContext = this;
            parentWindow = parent;
            partsToRequest = parts;

            updateAllPartsList();
            updateCurrentRequestParts();
        }

        /// <summary>
        /// Обновляет список с доступными запчастями.
        /// </summary>
        private void updateAllPartsList() {
            List<Part> allParts = DBEntities.Instance.Parts.ToList();

            allAvailableParts.ItemsSource = null;
            allAvailableParts.ItemsSource = allParts.Where(x => !partsToRequest.Select(c => c.Part).Contains(x));
        }

        /// <summary>
        /// Обновляет список с текущими выбранными запчастями.
        /// </summary>
        private void updateCurrentRequestParts() {
            currentRequestParts.ItemsSource = null;
            currentRequestParts.ItemsSource = partsToRequest;

            currentRequestParts.SelectedIndex = -1;
        }
        #endregion

        #region Функции обновления запчастей или их количества.

        /// <summary>
        /// Удаляет запчасть из списка запчастей для выполнения заказа.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void removePartFromRequest_Click(object sender, RoutedEventArgs e) {
            if (currentRequestParts.SelectedItem != null) {
                partsToRemove.Add(partsToRequest.Where(x => x.Part.Id == (currentRequestParts.SelectedItem as Parts_To_Request).Part.Id).FirstOrDefault());
                partsToRequest.Remove(partsToRequest.Where(x => x.Part.Id == (currentRequestParts.SelectedItem as Parts_To_Request).Part.Id).FirstOrDefault());

                updateAllPartsList();
                updateCurrentRequestParts();
            } else {
                MessageBox.Show("Не выбран элемент списка для удаления.", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Добавляет запчасть в список запчастей для выполнения заказа.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void addPartToRequest_Click(object sender, RoutedEventArgs e) {
            if (allAvailableParts.SelectedItem != null) {
                // Полностью инициализируем объект. Со всеми его свойствами.
                partsToRequest.Add(new Parts_To_Request() { 
                    Part = allAvailableParts.SelectedItem as Part,
                    Count = 1,
                    Part_Id = (allAvailableParts.SelectedItem as Part).Id,
                    Request_Id = parentWindow.request.Id,
                    Service_Request = parentWindow.request
                });

                updateAllPartsList();
                updateCurrentRequestParts();
            } else {
                MessageBox.Show("Не выбран элемент для добавления.", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Происходит при смене выбранного элемента в списке текущих запчастей.
        /// <br/>
        /// Проявляет элементы редактирования количества и настраивает их на текущую выбранную запчасть.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void currentRequestParts_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) {
            var selected = currentRequestParts.SelectedItem;
            
            if (selected != null) {
                partsCountDesc.Visibility = Visibility.Visible;
                partsCount.Visibility = Visibility.Visible;

                partsCountDesc.Text = $"Количество ({(selected as Parts_To_Request)?.Part.Part_Name}):";
                partsCount.Text = (selected as Parts_To_Request)?.Count.ToString();
            } else {
                partsCountDesc.Visibility = Visibility.Hidden;
                partsCount.Visibility = Visibility.Hidden;
            }
        }

        /// <summary>
        /// Происходит при смене текста в поле ввода количества запчастей.
        /// <br/>
        /// Регулирует количество у текущей выбранной запчасти.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void partsCount_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e) {
            var selected = currentRequestParts.SelectedItem as Parts_To_Request;

            if (selected != null) {
                int count = getSafePartCountAndUpdateTextBoxIfIncorrect();

                try {
                    partsToRequest.First(x => x.Part.Id == selected.Part.Id).Count = count;
                } catch (Exception ex) {
                    MessageBox.Show($"Произошла ошибка при обновлении количества запчастей.\n\n{ex.Message}.", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        /// <summary>
        /// Проверяет введенное количество на корректность.
        /// <br/>
        /// Если результат некорректен, значение в текстовом поле будет сброшено до "1".
        /// </summary>
        /// <returns>Безопасное значение количества.</returns>
        private int getSafePartCountAndUpdateTextBoxIfIncorrect() {
            string currentValue = partsCount.Text;

            if (!int.TryParse(currentValue, out int value)) {
                value = 1;

                partsCount.Text = value.ToString();
                MessageBox.Show("Введено некорректное значение для количества.\n\nЗначение сброшено.", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return value;
        }
        #endregion

        #region Функции кнопок.

        /// <summary>
        /// Позволяет создать новую деталь.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void createNewPart_Click(object sender, RoutedEventArgs e) {
            CreateNewServicePartWindow window = new CreateNewServicePartWindow();
            var result = window.ShowDialog();

            if (result.HasValue && result.Value) {
                DBEntities.Instance.Parts.Add(window.newPart);
                DBEntities.Instance.SaveChanges();

                updateAllPartsList();
            }
        }

        /// <summary>
        /// Сохраняет текущие детали заказа.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void saveChanges_Click(object sender, RoutedEventArgs e) {
            insertTethers();
            removeUnusedTethers();
            DialogResult = true;

            Close();
        }
        #endregion

        #region Функции обновления данных в БД.

        /// <summary>
        /// Вставляет в БД связи между заказом и запчастями.
        /// <br/>
        /// Это нужно, поскольку новые связи формируются без привязки к экземпляру контекста БД и автоматически не помещаются в БД.
        /// </summary>
        private void insertTethers() {
            var allRequests = DBEntities.Instance.Parts_To_Request.ToList();
            var newRequests = partsToRequest.Where(x => !allRequests.Contains(x));

            DBEntities.Instance.Parts_To_Request.AddRange(newRequests);
        }

        /// <summary>
        /// Удаляет старые и неактуальные связи между заказом и запчастями.
        /// <br/>
        /// Так как <i>"EF Core"</i> не умеет удалять их автоматически, значения ключей сбросятся в "null", что приведет к выбросу исключения при попытке сохранить данные.
        /// </summary>
        private void removeUnusedTethers() {
            try {
                DBEntities.Instance.Parts_To_Request.RemoveRange(partsToRemove);
            } catch {
                // Ничего не делаем. Просто если при удалении целого списка ни один элемент не будет найден, это выбросит исключение.
            }
        }
        #endregion
    }
}
