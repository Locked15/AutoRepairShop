using System;
using System.Linq;
using System.Collections.Generic;
using Auto_Repair_Shop.Entities;

namespace Auto_Repair_Shop.Classes {

    /// <summary>
    /// Новый уровень абстракции над классами формирования отчёта.
    /// </summary>
    public abstract class KirovReporting {

        /// <summary>
        /// Полное имя файла.
        /// </summary>
        public string fullFileName { get; set; }

        /// <summary>
        /// Использовать устаревший формат документа.
        /// </summary>
        public bool legacyDocumentFormat { get; set; }

        /// <summary>
        /// Список с заказами для формирования отчёта.
        /// </summary>
        public List<Service_Request> requests { get; set; }

        /// <summary>
        /// Конструктор класса по умолчанию. Абстрактного класса.
        /// </summary>
        public KirovReporting() {
            fullFileName = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            legacyDocumentFormat = false;

            requests = Enumerable.Empty<Service_Request>().ToList();
        }

        /// <summary>
        /// Конструктор класса. Абстрактного.
        /// </summary>
        /// <param name="fullFileName">Полный путь к файлу.</param>
        /// <param name="legacyDocumentFormat">Использовать устаревший формат документа.</param>
        /// <param name="requests">Список заказов для формирования отчёта.</param>
        public KirovReporting(string fullFileName, bool legacyDocumentFormat, List<Service_Request> requests) {
            this.fullFileName = fullFileName;
            this.legacyDocumentFormat = legacyDocumentFormat;
            this.requests = requests;
        }
    }
}
