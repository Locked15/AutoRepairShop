using System.Collections.Generic;
using Auto_Repair_Shop.Entities;

namespace Auto_Repair_Shop.Classes.Reporting { 

    /// <summary>
    /// Класс, нужный для формирования отчёта в Excel.
    /// <br/>
    /// В данном файле представлены базовые конструкции класса (конструкторы).
    /// </summary>
    public partial class ExcelReporting : KirovReporting {

        /// <summary>
        /// Конструктор класса.
        /// </summary>
        public ExcelReporting() : base() { }

        /// <summary>
        /// Конструктор класса.
        /// </summary>
        /// <param name="fullFileName">Полное название файла, в котором будет сохранен отчёт.</param>
        /// <param name="fontName">Шрифт для формирования отчёта.</param>
        /// <param name="legacyType">Использовать старый тип документа?</param>
        /// <param name="requests">Заказы, которые необходимо вставить в отчёт.</param>
        public ExcelReporting(string fullFileName, string fontName, bool legacyType, List<Service_Request> requests) : base(fullFileName, fontName, legacyType, requests) { }

        /// <summary>
        /// Начинает формирование отчёта в Excel.
        /// </summary>
        /// <returns>Успех формирования отчёта.</returns>
        public bool generateReport() {
            try {     
                if (legacyDocumentFormat) {
                    generateLegacyExcelReport();

                    return true;
                } else {
                    return false;
                }
            } catch {
                return false;
            }
        }
    }
}
