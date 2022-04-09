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
        protected string fullFileName { get; set; }

        /// <summary>
        /// Используемый шрифт для генерации документа.
        /// </summary>
        public string fontFamily { get; set; }

        /// <summary>
        /// Использовать устаревший формат документа.
        /// </summary>
        protected bool legacyDocumentFormat { get; set; }

        /// <summary>
        /// Список с заказами для формирования отчёта.
        /// </summary>
        protected List<Service_Request> requests { get; set; }

        /// <summary>
        /// Конструктор класса по умолчанию. Абстрактного класса.
        /// </summary>
        protected KirovReporting() {
            fullFileName = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            legacyDocumentFormat = false;
            fontFamily = "Georgia";

            requests = Enumerable.Empty<Service_Request>().ToList();
        }

        /// <summary>
        /// Конструктор класса. Абстрактного.
        /// </summary>
        /// <param name="fullFileName">Полный путь к файлу.</param>
        /// <param name="legacyDocumentFormat">Использовать устаревший формат документа.</param>
        /// <param name="requests">Список заказов для формирования отчёта.</param>
        protected KirovReporting(string fullFileName, string fontFamily, bool legacyDocumentFormat, List<Service_Request> requests) {
            this.fullFileName = fullFileName;
            this.legacyDocumentFormat = legacyDocumentFormat;
            this.fontFamily = fontFamily;
            this.requests = requests;

            if (!ProgramSettings.settings.showCompletedRequests) {
                requests = requests.Where(x => x.Request_Approx_Complete > DateTime.Now).ToList();
            }
        }
    }
}
