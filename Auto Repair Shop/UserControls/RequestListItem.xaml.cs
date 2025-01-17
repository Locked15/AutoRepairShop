﻿using System;
using System.Linq;
using System.Windows.Media;
using System.Windows.Controls;
using Auto_Repair_Shop.Entities;
using Auto_Repair_Shop.Resources;

namespace Auto_Repair_Shop.UserControls { 

    /// <summary>
    /// Элемент списка заказов.
    /// </summary>
    public partial class RequestListItem : UserControl { 

        /// <summary>
        /// Свойство, содержащее заказ, относящийся к данному элементу.
        /// </summary>
        public Service_Request request { get; set; }

        #region Функции инициализации.

        /// <summary>
        /// Конструктор класса.
        /// </summary>
        /// <param name="request">Заказ.</param>
        public RequestListItem(Service_Request request) {
            InitializeComponent();

            DataContext = request;
            this.request = request;

            bindImages();
            initializeStateNumber();
            initializePrice();
            insertParts();
            updateColor();
        }

        /// <summary>
        /// Так как путь изображений относителен проекта, то привязывать изображения придется вручную.
        /// </summary>
        private void bindImages() {
            vehicleImage.Source = new System.Windows.Media.Imaging.BitmapImage(
                                  new Uri(ResourceManager.checkExistsAndReturnFullPath(request.Vehicle.Image)));

            brandImage.Source = new System.Windows.Media.Imaging.BitmapImage(
                                new Uri(ResourceManager.checkExistsAndReturnFullPath(request.Vehicle.Vehicle_Brand.Image)));
        }

        /// <summary>
        /// Вставляет в элемент государственный номер автомобиля.
        /// </summary>
        private void initializeStateNumber() => vehicleStateNumber.Text = request.Vehicle.getDividedStateNumber();

        /// <summary>
        /// Рассчитывает стоимость заказа на основе стоимости комплектующих.
        /// <br/>
        /// Стоимость рассчитывается так:
        /// 1. Учитывается полная стоимость всех деталей и их количества;
        /// 2. Потом она умножается на базовую стоимость услуги;
        /// 3. Затем она перемножается с модификатором стоимости, зависящим от класса автомобиля.
        /// </summary>
        private void initializePrice() {
            decimal price = request.calculateTotalPrice();

            requestPrice.Text = price > 0 ? price.ToString(".00") : "бесплатно";
        }

        /// <summary>
        /// Вставляет в визуальный элемент все нужные запчасти.
        /// </summary>
        private void insertParts() {
            string parts = string.Empty;

            foreach (var item in request.Parts_To_Request) {
                parts += $"{item.Part.Part_Name} (кол-во: {item.Count}){(item == request.Parts_To_Request.Last() ? "." : ", ")}";
            }

            requestParts.Text = string.IsNullOrEmpty(parts) ? "Нет запчастей." : parts;
        }

        /// <summary>
        /// Обновляет цвет заднего фона в зависимости от значений свойств.
        /// </summary>
        private void updateColor() {
            string color = request.Vehicle.getColorFromClass();

            Background = ColorConverter.ConvertFromString(color) as Brush;
        }
        #endregion
    }
}
