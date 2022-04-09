using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using NPOI.XWPF.UserModel;
using Auto_Repair_Shop.Entities;
using Auto_Repair_Shop.Resources;

namespace Auto_Repair_Shop.Classes.Reporting {

    /// <summary>
    /// Класс, нужный для формирования отчёта.
    /// </summary>
    public class WordReporting : KirovReporting {

        #region Конструкторы класса.

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
        public WordReporting(string fullFileName, string fontFamily, bool legacyDocumentFormat, List<Service_Request> requests) : base(fullFileName, fontFamily, legacyDocumentFormat, requests) { }
        #endregion

        #region Функции.

        /// <summary>
        /// Формирует отчётность.
        /// </summary>
        /// <returns>Удалось ли сформировать отчётность.</returns>
        public bool generateReport() {
            try {
                generateActualReport();

                return true;
            } catch {
                return false;
            }
        }

        /// <summary>
        /// Формирует отчёт современного типа (расширение ".docx").
        /// </summary>
        /// <exception cref="IOException"/>
        private void generateActualReport() {
            XWPFDocument document = new XWPFDocument();

            foreach (var request in requests) {
                generateHeaders(document, request);
                generateTable(document, request);
                generatePageBreak(document);
            }

            using (FileStream fs = new FileStream(fullFileName, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)) {
                document.Write(fs);
            }

            document.Close();
        }

        #region Формирование базовых элементов.

        /// <summary>
        /// Генерирует параграф заголовка страницы.
        /// </summary>
        /// <param name="document">Документ для генерации заголовка.</param>
        /// <param name="request">Информация о заказе.</param>
        private void generateHeaders(XWPFDocument document, Service_Request request) {
            XWPFParagraph paragraph = document.CreateParagraph();
            paragraph.Alignment = ParagraphAlignment.CENTER;

            XWPFRun run = paragraph.CreateRun();
            run.FontSize = 26;
            run.SetFontFamily(fontFamily, FontCharRange.None);
            run.SetText(request.Service_Type.Name.TrimEnd('.'));
            run.AddBreak(BreakType.TEXTWRAPPING);
        }

        /// <summary>
        /// Создает таблицу с информацией о заказе.
        /// </summary>
        /// <param name="document">Документ, в котором будет создана таблица.</param>
        /// <param name="request">Запрос, который нужно поместить в таблицу.</param>
        private void generateTable(XWPFDocument document, Service_Request request) {
            XWPFTable table = document.CreateTable(4, 2);

            generateVehicleNameRow_First(table, request);
            generateVehicleImageRow_Second(table, request);
            generateSideInformationRow_Third(table, request);
            generateServicePartsRow_Fourth(table, request);
        }
        #endregion

        #region Формирование строк с основной информацией.

        /// <summary>
        /// Заполняет содержимым первую строку таблицы, помещая туда информацию о бренде и названии машины.
        /// </summary>
        /// <param name="table">Таблица, в которую будут внесены значения.</param>
        /// <param name="request">Запрос, который нужно внести.</param>
        private void generateVehicleNameRow_First(XWPFTable table, Service_Request request) {
            XWPFTableRow row = table.GetRow(0);
            row.MergeCells(0, 1);

            XWPFTableCell cell = row.GetCell(0);
            XWPFParagraph paragraph = cell.AddParagraph();
            paragraph.Alignment = ParagraphAlignment.CENTER;

            XWPFRun run = paragraph.CreateRun();
            run.FontSize = 18;

            run.SetFontFamily(fontFamily, FontCharRange.None);
            run.SetText($"{request.Vehicle.Name} — {request.Vehicle.Vehicle_Brand.Brand}");
            run.SetColor(request.Vehicle.getColorFromClass());
            run.AddBreak(BreakType.TEXTWRAPPING);
        }

        /// <summary>
        /// Заполняет содержимым вторую строку таблицы, помещая туда изображение автомобиля.
        /// </summary>
        /// <param name="table">Таблица, в которую будут внесены значения.</param>
        /// <param name="request">Запрос, который нужно внести.</param>
        private void generateVehicleImageRow_Second(XWPFTable table, Service_Request request) {
            XWPFTableRow row = table.GetRow(1);
            row.MergeCells(0, 1);

            XWPFTableCell cell = row.GetCell(0);
            XWPFParagraph paragraph = cell.AddParagraph();
            paragraph.Alignment= ParagraphAlignment.CENTER;

            XWPFRun run = paragraph.CreateRun();

            using (FileStream stream = new FileStream(ResourceManager.checkExistsAndReturnFullPath(request.Vehicle.Image), FileMode.Open, FileAccess.Read)) {
                XWPFPicture picture = run.AddPicture(stream, (int)PictureType.PNG, "Машина", 600 * 12857, 300 * 12857);
            }
        }

        /// <summary>
        /// Заполняет содержимым третью строку таблицы, помещая туда информацию о датах заказа и его стоимости.
        /// </summary>
        /// <param name="table">Таблица, в которую будут внесены значения.</param>
        /// <param name="request">Запрос, который нужно внести.</param>
        private void generateSideInformationRow_Third(XWPFTable table, Service_Request request) {
            XWPFTableRow row = table.GetRow(2);
            generateDatesCell(row, request.Request_Date, request.Request_Approx_Complete);
            generatePriceCell(row, request.calculateTotalPrice());
        }

        #region Формирование ячеек.

        /// <summary>
        /// Заполняет содержимым первую ячейку третьей строки, помещая туда данные о дате размещения заказа и дате выполнения.
        /// </summary>
        /// <param name="row">Строка, в которой находится нужная ячейка.</param>
        /// <param name="begin">Дата размещения заказа.</param>
        /// <param name="end">Дата выполнения заказа.</param>
        private void generateDatesCell(XWPFTableRow row, DateTime? begin, DateTime? end) {
            XWPFTableCell datesCell = row.GetCell(0);

            // Дата начала:
            XWPFParagraph paragraph = datesCell.AddParagraph();
            paragraph.Alignment = ParagraphAlignment.CENTER;

            XWPFRun run = paragraph.CreateRun();
            run.SetText($"Дата начала: {(begin.HasValue ? begin.Value.ToString("dd.MM.yyyy") : "Неизвестно")}");
            run.SetFontFamily(fontFamily, FontCharRange.None);
            run.FontSize = 15;

            // Дата конца:
            XWPFParagraph secondParagraph = datesCell.AddParagraph();
            secondParagraph.Alignment = ParagraphAlignment.CENTER;

            XWPFRun secondRun = secondParagraph.CreateRun();
            secondRun.SetText($"Дата конца: {(end.HasValue ? end.Value.ToString("dd.MM.yyyy") : "Неизвестно")}");
            secondRun.SetFontFamily(fontFamily, FontCharRange.None);
            secondRun.FontSize = 15;

            secondRun.AddBreak(BreakType.TEXTWRAPPING);
        }

        /// <summary>
        /// Заполняет содержимым вторую ячейку третьей строки, помещая туда данные о стоимости заказа.
        /// </summary>
        /// <param name="row">Строка, в которой находится нужная ячейка.</param>
        /// <param name="price">Стоимость заказа.</param>
        private void generatePriceCell(XWPFTableRow row, decimal price) {
            XWPFTableCell priceCell = row.GetCell(1);
            XWPFParagraph priceParagraph = priceCell.AddParagraph();
            priceParagraph.Alignment = ParagraphAlignment.CENTER;

            XWPFRun run = priceParagraph.CreateRun();
            run.SetFontFamily(fontFamily, FontCharRange.None);
            run.SetText($"Итоговая стоимость заказа: {(price > 0 ? price.ToString(".00") : "бесплатно")}.");
            run.FontSize = 16;
        }
        #endregion
        #endregion

        #region Формирование заключительных частей страницы.

        /// <summary>
        /// Заполняет содержимым четвертую строку таблицы, помещая туда информацию о запчастях, нужных для выполнения заказа.
        /// </summary>
        /// <param name="table">Таблица, в которую нужно внести данные.</param>
        /// <param name="request">Заказ, который нужно разместить в таблице.</param>
        private void generateServicePartsRow_Fourth(XWPFTable table, Service_Request request) {
            XWPFTableRow row = table.GetRow(3);
            row.MergeCells(0, 1);

            XWPFTableCell cell = row.GetCell(0);
            XWPFParagraph paragraph = cell.AddParagraph();
            XWPFRun run = paragraph.CreateRun();

            run.SetFontFamily(fontFamily, FontCharRange.None);
            run.FontSize = 16;
            run.SetText("Детали обслуживания: ");

            if (request.Parts_To_Request.Count == 0) {
                run.AppendText("отсутствуют.");
            } else {
                foreach (var item in request.Parts_To_Request) {
                    string endOfLine = item.Equals(request.Parts_To_Request.LastOrDefault()) ? "." : ", ";

                    run.AppendText($"{item.Part.Part_Name} (кол-во: {item.Count}){endOfLine}");
                }
            }

            run.AddBreak(BreakType.TEXTWRAPPING);
        }

        /// <summary>
        /// Вставляет разрыв страницы после завершения таблицы заказа.
        /// </summary>
        /// <param name="document">Документ, где будет вставлен разрыв.</param>
        private void generatePageBreak(XWPFDocument document) {
            XWPFParagraph paragraph = document.CreateParagraph();
            XWPFRun run = paragraph.CreateRun();

            run.AddBreak(BreakType.PAGE);
        }
        #endregion
        #endregion
    }
}
