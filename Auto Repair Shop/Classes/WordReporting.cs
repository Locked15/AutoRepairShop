using Auto_Repair_Shop.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Auto_Repair_Shop.Classes {

    /// <summary>
    /// Класс, нужный для формирования отчёта.
    /// </summary>
    public class WordReporting : KirovReporting {

        /// <summary>
        /// Конструктор класса.
        /// </summary>
        public WordReporting() : base() { }

        /// <summary>
        /// Конструктор класса.
        /// </summary>
        /// <param name="fullFileName">Полное имя файла.</param>
        /// <param name="legacyDocumentFormat">Использовать устаревший формат документа.</param>
        /// <param name="requests">Список заказов для формирования отчёта.</param>
        public WordReporting(string fullFileName, bool legacyDocumentFormat, List<Service_Request> requests) : base(fullFileName, legacyDocumentFormat, requests) { }

        /// <summary>
        /// Формирует отчётность.
        /// </summary>
        /// <returns>Удалось ли сформировать отчётность.</returns>
        public bool generateReport() {
            bool result;

            if (legacyDocumentFormat) {
                try {
                    generateLegacyReport();

                    result = true;
                } catch {
                    result = false;
                }
            } else {
                try {
                    generateActualReport();

                    result = true;
                } catch {
                    result = false;
                }
            }

            return result;
        }

        #region Формирование отчёта устаревшего типа.

        /// <summary>
        /// Формирует отчёт устаревшего типа (расширение ".doc").
        /// </summary>
        /// <returns>Успех формирования отчёта.</returns>
        private bool generateLegacyReport() {
            return false;
        }
        #endregion

        #region Формирование отчёта современного типа.

        /// <summary>
        /// Формирует отчёт современного типа (расширение ".docx").
        /// </summary>
        /// <returns>Успех формирования отчёта.</returns>
        private bool generateActualReport() {
            return false;
        }
        #endregion
    }
}
