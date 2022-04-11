using System;
using System.Linq;
using System.Windows;
using System.Collections.Generic;
using Auto_Repair_Shop.Classes;
using Auto_Repair_Shop.Entities;
using Auto_Repair_Shop.UserControls;
using Auto_Repair_Shop.Windows.CreatingWindows;

namespace Auto_Repair_Shop.Windows {

    /// <summary>
    /// Логика главного окна, выводящего текущие заказы. 
    /// </summary>
    public partial class MainWindow : Window {

        /// <summary>
        /// Свойство, содержащее все запросы обслуживания, соответствующие выборке.
        /// </summary>
        public List<Service_Request> selectedRequests { get; set; } = DBEntities.Instance.Service_Request.ToList();

        #region Функции инициализации.

        /// <summary>
        /// Конструктор класса. 
        /// </summary>
        public MainWindow() {
            InitializeComponent();
            initializeRequestTypes();
            initializeListeners();

            selectionChanged(default, default);
        }

        /// <summary>
        /// Устанавливает обработчики событий для элементов окна.
        /// </summary>
        private void initializeListeners() {
            searchBox.TextChanged += selectionChanged;
            sortBox.SelectionChanged += selectionChanged;
            filterBox.SelectionChanged += selectionChanged;

            pageSelector.SelectionChanged += pageSelector_SelectionChanged;
        }

        /// <summary>
        /// Заполняет данными список доступных вариантов фильтрации.
        /// </summary>
        private void initializeRequestTypes() {
            var types = new List<string>(1) {
                "Все типы"
            };
            types.AddRange(DBEntities.Instance.Service_Type.Select(x => x.Name));

            filterBox.Items.Clear();
            filterBox.ItemsSource = types;
            filterBox.SelectedIndex = 0;
        }
        #endregion

        #region Функции выборки.

        /// <summary>
        /// Происходит при смене какой-либо функции выборки (поиск, сортировка или фильтрация).
        /// <br/>
        /// Обновляет список доступных заказов.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void selectionChanged(object sender, dynamic e) {
            var search = searchBox.Text.ToLower();
            selectedRequests = DBEntities.Instance.Service_Request.ToList();

            if (!ProgramSettings.settings.showCompletedRequests) {
                selectedRequests = selectedRequests.Where(x => x.Request_Approx_Complete.HasValue && x.Request_Approx_Complete.Value < DateTime.Now).ToList();
            }

            selectedRequests = selectedRequests.Where(x => x.Vehicle.Name.ToLower().Contains(search) || 
                                                      x.Service_Type.Name.ToLower().Contains(search) || 
                                                     (x.Request_Approx_Complete.HasValue && x.Request_Approx_Complete.Value.ToString("dd.MM.yyyy").Contains(search))).ToList();

            if (sortBox.SelectedIndex > 0) {
                if (sortBox.SelectedIndex < 5) {
                    switch (sortBox.SelectedIndex) {
                        case 2:
                            selectedRequests = selectedRequests.OrderBy(x => x.calculateTotalPrice()).ToList();

                            break; 
                        case 3:
                            selectedRequests = selectedRequests.OrderBy(x => x.Parts_To_Request.Sum(y => y.Count)).ToList();

                            break;
                        case 4:
                            selectedRequests = selectedRequests.OrderBy(x => x.Request_Approx_Complete).ToList();

                            break;
                    }
                } else {
                    switch (sortBox.SelectedIndex)
                    {
                        case 6:
                            selectedRequests = selectedRequests.OrderByDescending(x => x.calculateTotalPrice()).ToList();

                            break;
                        case 7:
                            selectedRequests = selectedRequests.OrderByDescending(x => x.Parts_To_Request.Sum(y => y.Count)).ToList();

                            break;
                        case 8:
                            selectedRequests = selectedRequests.OrderByDescending(x => x.Request_Approx_Complete).ToList();

                            break;
                    }
                }
            }

            if (filterBox.SelectedIndex > 0) {
                var selected = filterBox.SelectedItem as string;

                selectedRequests = selectedRequests.Where(x => x.Service_Type.Name.Equals(selected)).ToList();
            }

            updatePages();

            if (selectedRequests.Count == 0) {
                MessageBox.Show("По заданным параметрам заказов не найдено.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        /// <summary>
        /// Обновляет список доступных страниц. 
        /// <br/>
        /// Сбрасывает выбранную страницу и вставляет товары заново.
        /// </summary>
        private void updatePages() {
            int pages = selectedRequests.Count / 20;

            if (selectedRequests.Count % 20 > 0) {
                ++pages;
            }

            if (pages == 0) {
                pages = 1;
            }

            pageSelector.ItemsSource = Enumerable.Range(1, pages);
            if (pageSelector.SelectedIndex == 0) {
                pageSelector_SelectionChanged(default, default);
            } else {
                pageSelector.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// Вставляет заказы в список на окне.
        /// </summary>
        /// <param name="requests">Список заказов для вставки.</param>
        private void insertRequestsToList(List<Service_Request> requests) {
            double width;
            activeRequests.Items.Clear();

            foreach (var request in requests) {
                if (WindowState == WindowState.Maximized) {
                    width = RenderSize.Width - 45;
                } else {
                    width = Width - 45;
                }

                var item = new RequestListItem(request) {
                    Width = width
                };

                activeRequests.Items.Add(item);
            }
        }
        #endregion

        #region Функции списка заказов.

        /// <summary>
        /// Возникает при двойном нажатии на элемент списка заказов.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void activeRequests_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e) {
            var selectedItem = (activeRequests.SelectedItem as RequestListItem)?.request;

            // Если пользователь дважды кликнет на пустой области, произойдет добавление товара. Чтобы этого избежать, делаем проверку.
            if (selectedItem != null) {
                var window = new RequestSettingWindow(selectedItem);
                var result = window.ShowDialog();

                if (result.HasValue && result.Value) {
                    if (window.toDelete) {
                        if (selectedItem.Parts_To_Request.Count > 0)
                            DBEntities.Instance.Parts_To_Request.RemoveRange(selectedItem.Parts_To_Request);
                        DBEntities.Instance.Service_Request.Remove(selectedItem);
                    }

                    try {
                        DBEntities.Instance.SaveChanges();

                        selectionChanged(default, default);
                    } catch {
                        MessageBox.Show("Каждый раз, когда тестировщик пытается сломать программу, в мире грустит один программист." +
                                        "\n\nНе нужно печалить программистов — не ломайте программы.", 
                                        "Манифест", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }
            }
        }
        #endregion

        #region Функции нижней панели.

        /// <summary>
        /// Открывает окно настроек программы.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void settings_Click(object sender, RoutedEventArgs e) {
            SettingsWindow window = new SettingsWindow();
            bool? result = window.ShowDialog();

            if (result.HasValue && result.Value) {
                selectionChanged(default, default);
            }
        }

        /// <summary>
        /// Открывает окно формирования отчётности.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void serviceReporting_Click(object sender, RoutedEventArgs e) {
            ReportingWindow window = new ReportingWindow();
            window.ShowDialog();
        }

        /// <summary>
        /// Открывает окно добавления нового заказа.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void addServiceRequest_Click(object sender, RoutedEventArgs e) {
            var window = new RequestSettingWindow(null);
            var result = window.ShowDialog();

            if (result.HasValue && result.Value) {
                DBEntities.Instance.Service_Request.Add(window.request);
                DBEntities.Instance.SaveChanges();

                selectionChanged(default, default);
            }
        }
        #endregion

        #region Функции навигации.

        /// <summary>
        /// Переводит пользователя на предыдущую страницу.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void previousPage_Click(object sender, RoutedEventArgs e) {
            if (pageSelector.SelectedIndex > 0) {
                --pageSelector.SelectedIndex;
            }
        }

        /// <summary>
        /// Происходит при смене выбранной страницы.
        /// <br/>
        /// Обновляет выведенные на экран заказы.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void pageSelector_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) {
            var current = selectedRequests.Skip(pageSelector.SelectedIndex * 20).Take(20);

            insertRequestsToList(current.ToList());
        }

        /// <summary>
        /// Переводит пользователя на следующую страницу.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void nextPage_Click(object sender, RoutedEventArgs e) {
            if (pageSelector.SelectedIndex + 1 < pageSelector.Items.Count) {
                ++pageSelector.SelectedIndex;
            }
        }
        #endregion

        #region Прочие события.

        /// <summary>
        /// Происходит при изменении размера окна.
        /// <br/>
        /// Подстраивает все выведенные заказы под новый размер.
        /// </summary>
        /// <param name="sender">Объект, вызвавший событие.</param>
        /// <param name="e">Аргументы события.</param>
        private void window_SizeChanged(object sender, SizeChangedEventArgs e) {
            double width;
            try {
                width = RenderSize.Width - 45;
            } catch {
                width = Width - 45;
            }

            foreach (RequestListItem item in activeRequests.Items) {
                item.Width = width;
            }
        }
        #endregion
    }
}
