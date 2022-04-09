using System.Windows.Media;
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
                } else {
                   generateModernExcelReport();
                }

                return true;
            } catch {
                return false;
            }
        }

        /// <summary>
        /// Конвертирует строковое представление цвета (HEX) в массив байтов (RGB).
        /// </summary>
        /// <param name="hex">Цвет в HEX-формате.</param>
        /// <returns>Массив байтов в формате RGB.</returns>
        private byte[] convertFromHexToRgb(string hex)
        {
            SolidColorBrush brush = new BrushConverter().ConvertFromString(hex) as SolidColorBrush;
            Color color = brush.Color;

            return new byte[] { color.R, color.G, color.B };
        }
    }
}
