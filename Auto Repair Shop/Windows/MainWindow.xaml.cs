using System;
using System.Linq;
using System.Windows;
using System.Collections.Generic;
using Auto_Repair_Shop.Classes;
using Auto_Repair_Shop.Entities;
using Auto_Repair_Shop.UserControls;

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

        private void initializeListeners() {
            searchBox.TextChanged += selectionChanged;
            sortBox.SelectionChanged += selectionChanged;
            filterBox.SelectionChanged += selectionChanged;

            pageSelector.SelectionChanged += pageSelector_SelectionChanged;
        }

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

        private void insertRequestsToList(List<Service_Request> requests) {
            ActiveRequests.Items.Clear();

            foreach (var request in requests) {
                var item = new RequestListItem(request) {
                    Width = Width - 45
                };

                ActiveRequests.Items.Add(item);
            }
        }
        #endregion

        #region Функции списка заказов.
        private void activeRequests_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e) {

        }

        #endregion

        #region Функции нижней панели.
        private void settings_Click(object sender, RoutedEventArgs e) {
            SettingsWindow window = new SettingsWindow();

            window.ShowDialog();
        }

        private void serviceReporting_Click(object sender, RoutedEventArgs e) {

        }

        private void addServiceRequest_Click(object sender, RoutedEventArgs e) {

        }
        #endregion

        #region Функции навигации.
        private void previousPage_Click(object sender, RoutedEventArgs e) {
            if (pageSelector.SelectedIndex > 0) {
                --pageSelector.SelectedIndex;
            }
        }

        private void pageSelector_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) {
            var current = selectedRequests.Skip(pageSelector.SelectedIndex * 20).Take(20);

            insertRequestsToList(current.ToList());
        }

        private void nextPage_Click(object sender, RoutedEventArgs e) {
            if (pageSelector.SelectedIndex + 1 < pageSelector.Items.Count) {
                ++pageSelector.SelectedIndex;
            }
        }
        #endregion

        #region Прочие события.
        private void window_SizeChanged(object sender, SizeChangedEventArgs e) {
            double width;
            try {
                width = RenderSize.Width - 45;
            } catch {
                width = Width - 45;
            }

            foreach (RequestListItem item in ActiveRequests.Items) {
                item.Width = width;
            }
        }
        #endregion
    }
}
