using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Auto_Repair_Shop.Entities;
using Auto_Repair_Shop.Resources;

namespace Auto_Repair_Shop.Classes.Reporting {
    
    /// <summary>
    /// Класс формирования отчёта в Excel.
    /// <br/>
    /// Данный файл содержит функционал формирования отчёта в современной версии Excel.
    /// </summary>
    public partial class ExcelReporting {
        
        /// <summary>
        /// Начинает формирование отчёта в современной версии Excel.
        /// </summary>
        private void generateModernExcelReport() {
            XSSFWorkbook document = new XSSFWorkbook();

            foreach (var request in requests) {
                XSSFSheet sheet = (XSSFSheet)document.CreateSheet(getSafeSheetName_Modern(document, request));
                insertSheetRowsAndCells_First_Modern(sheet);

                insertServiceInformationRow_Second_Modern(document, sheet, request);
                insertVehicleInformationRow_Third_Modern(document, sheet, request);
                insertVehicleImageRow_Fourth_Modern(document, sheet, request);
                insertRequestDatesRow_Fifth_Modern(document, sheet, request);
                insertRequestPrice_Sixth_Modern(document, sheet, request);
                insertRequestPartsRows_Seventh_Modern(document, sheet, request);
            }

            using (FileStream fs = new FileStream(fullFileName, FileMode.Create, FileAccess.ReadWrite, FileShare.Read)) {
                document.Write(fs);
            }

            document.Close();
        }

        /// <summary>
        /// Проверяет документ на существование страницы с именем, которое будет у новой страницы.
        /// </summary>
        /// <param name="document">Документ для проверки.</param>
        /// <param name="request">Запрос, на основе которого строится имя новой страницы.</param>
        /// <returns>Оригинальное название страницы, если дубликатов нет. В ином случае добавляется окончание, чтобы избежать исключений.</returns>
        private string getSafeSheetName_Modern(XSSFWorkbook document, Service_Request request) {
            int error = 0;
            string sheetName = request.Vehicle.getDividedStateNumber();

            while (document.GetSheetIndex(sheetName) != -1) {
                sheetName = $"{request.Vehicle.getDividedStateNumber()} — {++error}";
            }

            return sheetName;
        }

        #region Формирование базовой информации.

        /// <summary>
        /// Формирует строки и ячейки в отчёте.
        /// <br/>
        /// Работает с современной версией Excel.
        /// </summary>
        /// <param name="sheet">Страница документа для формирования элементов.</param>
        private void insertSheetRowsAndCells_First_Modern(XSSFSheet sheet) {
            for (int i = 0; i < 15; i++) {
                XSSFRow row = (XSSFRow)sheet.CreateRow(i);
                row.Height = 500;

                for (int j = 0; j < 2; j++) {
                    row.CreateCell(j);
                }
            }

            sheet.SetColumnWidth(0, 7500);
            sheet.SetColumnWidth(1, 7500);
        }

        /// <summary>
        /// Формирует строку с информацией о типе заказа.
        /// <br/>
        /// Работает с современной версией Excel.
        /// </summary>
        /// <param name="document">Документ, в котором находится нужная страница.</param>
        /// <param name="sheet">Страница, на которой будет размещена информация.</param>
        /// <param name="request">Заказ, содержащий информацию.</param>
        private void insertServiceInformationRow_Second_Modern(XSSFWorkbook document, XSSFSheet sheet, Service_Request request) {
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, 1));

            XSSFCell mergedCell = (XSSFCell)sheet.GetRow(0).GetCell(0);
            mergedCell.CellStyle = generateCellStyle_Modern(document, true, true, 26);
            mergedCell.SetCellValue(request.Service_Type.Name.TrimEnd('.'));
        }

        /// <summary>
        /// Формирует строку с информацией о автомобиле (бренд и название).
        /// <br/>
        /// Работает с современный версией Excel.
        /// </summary>
        /// <param name="document">Документ, к которому относится таргетированная страница.</param>
        /// <param name="sheet">Страница, в которую будет встроена информация.</param>
        /// <param name="request">Заказ, содержащий информацию.</param>
        private void insertVehicleInformationRow_Third_Modern(XSSFWorkbook document, XSSFSheet sheet, Service_Request request) {
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(1, 1, 0, 1));

            XSSFCell mergedCell = (XSSFCell)sheet.GetRow(1).GetCell(0);
            mergedCell.CellStyle = generateCellStyle_Modern(document, true, true, 20, request.Vehicle.getColorFromClass());
            mergedCell.SetCellValue($"{request.Vehicle.Name} — {request.Vehicle.Vehicle_Brand.Brand}");
        }

        #region Добавление изображения.

        /// <summary>
        /// Добавляет изображение автомобиля в отчёт. Выравнивает его по якорным точкам.
        /// <br/>
        /// Работает с современной версией Excel.
        /// </summary>
        /// <param name="document">Документ, к которому относится таргетированная страница.</param>
        /// <param name="sheet">Страница, в которую будет встроена информация.</param>
        /// <param name="request">Заказ, содержащий информацию.</param>
        private void insertVehicleImageRow_Fourth_Modern(XSSFWorkbook document, XSSFSheet sheet, Service_Request request) {
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(2, 9, 0, 1));

            // Устанавливаем изменение размера для изображения.
            XSSFPicture picture = (XSSFPicture)createPicture_Modern(document, sheet, request);
            picture.Resize(1, 1);
        }

        /// <summary>
        /// Добавляет в документ изображение и инициализирует высокоуровневые объекты для работы с ним.
        /// <br/>
        /// Работает с современной версией Excel.
        /// </summary>
        /// <param name="document">Документ, в который будет добавлено изображение.</param>
        /// <param name="sheet">Страница, на которой будет размещено изображение.</param>
        /// <param name="request">Запрос, который содержит адрес изображения.</param>
        /// <returns>Высокоуровневый объект для работы с изображением.</returns>
        private IPicture createPicture_Modern(XSSFWorkbook document, XSSFSheet sheet, Service_Request request) {
            byte[] pictureMap = File.ReadAllBytes(ResourceManager.checkExistsAndReturnFullPath(request.Vehicle.Image));
            int pictureIndex = document.AddPicture(pictureMap, PictureType.PNG);

            ICreationHelper helper = document.GetCreationHelper();
            IDrawing patriarch = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = helper.CreateClientAnchor();

            /* Важно: Якорь берет границу не учитывая граничные значения.
               Иными словами, указав границу на 10 строке, последний фрагмент изображения будет находиться в 9 строке. */
            anchor.AnchorType = AnchorType.MoveAndResize;
            anchor.Row1 = 2;
            anchor.Col1 = 0;
            anchor.Row2 = 10;
            anchor.Col2 = 2;

            return patriarch.CreatePicture(anchor, pictureIndex);
        }
        #endregion
        #endregion

        #region Формирование вторичной информации.

        #region Добавление дат.

        /// <summary>
        /// Формирует строки с датами заказа (начало и конец).
        /// <br/>
        /// Работает с современной версией Excel.
        /// </summary>
        /// <param name="document">Документ, к которому принадлежит нужная страница.</param>
        /// <param name="sheet">Страница, в которую будет встроена информация.</param>
        /// <param name="request">Заказ.</param>
        private void insertRequestDatesRow_Fifth_Modern(XSSFWorkbook document, XSSFSheet sheet, Service_Request request) {
            generateBeginDate_Modern(document, sheet, request);

            generateCompleteDate_Modern(document, sheet, request);
        }

        /// <summary>
        /// Формирует на странице строку с датой размещения заказа.
        /// <br/>
        /// Работает с современной версией Excel.
        /// </summary>
        /// <param name="document">Документ, к которому принадлежит нужная страница.</param>
        /// <param name="sheet">Страница, в которую будет встроена информация.</param>
        /// <param name="request">Заказ.</param>
        private void generateBeginDate_Modern(XSSFWorkbook document, XSSFSheet sheet, Service_Request request) {
            XSSFCell descCell = (XSSFCell)sheet.GetRow(10).GetCell(0);
            descCell.CellStyle = generateCellStyle_Modern(document, false, true, 16);
            descCell.SetCellValue($"Дата размещения:");

            XSSFCell valueCell = (XSSFCell)sheet.GetRow(10).GetCell(1);
            valueCell.CellStyle = generateCellStyle_Modern(document, false, true, 16);
            valueCell.CellStyle.Alignment = HorizontalAlignment.Right;
            valueCell.SetCellValue($"{request.Request_Date.Value:dd.MM.yyyy}.");
        }

        /// <summary>
        /// Формирует на странице строку с датой выполнения заказа.
        /// <br/>
        /// Работает с современной версией Excel.
        /// </summary>
        /// <param name="document">Документ, к которому принадлежит нужная страница.</param>
        /// <param name="sheet">Страница, в которую будет встроена информация.</param>
        /// <param name="request">Заказ.</param>
        private void generateCompleteDate_Modern(XSSFWorkbook document, XSSFSheet sheet, Service_Request request) {
            XSSFCell descCell = (XSSFCell)sheet.GetRow(11).GetCell(0);
            descCell.CellStyle = generateCellStyle_Modern(document, false, true, 16);
            descCell.SetCellValue($"Дата выполнения:");

            XSSFCell valueCell = (XSSFCell)sheet.GetRow(11).GetCell(1);
            valueCell.CellStyle = generateCellStyle_Modern(document, false, true, 16);
            valueCell.CellStyle.Alignment = HorizontalAlignment.Right;
            valueCell.SetCellValue($"{request.Request_Approx_Complete.Value:dd.MM.yyyy}.");
        }
        #endregion

        /// <summary>
        /// Формирует на странице строку со стоимостью заказа.
        /// <br/>
        /// Работает с современной версией Excel.
        /// </summary>
        /// <param name="document">Документ, к которому принадлежит нужная страница.</param>
        /// <param name="sheet">Страница, в которую будет встроена информация.</param>
        /// <param name="request">Заказ.</param>
        private void insertRequestPrice_Sixth_Modern(XSSFWorkbook document, XSSFSheet sheet, Service_Request request) {
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(12, 12, 0, 1));
            decimal price = request.calculateTotalPrice();

            XSSFCell mergedCell = (XSSFCell)sheet.GetRow(12).GetCell(0);
            mergedCell.CellStyle = generateCellStyle_Modern(document, true, true, 22);
            mergedCell.SetCellValue($"Полная стоимость: {(price > 0 ? price.ToString(".00") : "бесплатно")}.");
        }

        #region Добавление запчастей, нужных для выполнения заказа.

        /// <summary>
        /// Формирует строку с информацией про необходимые для работы запчасти.
        /// <br/>
        /// Конкретно этот метод только добавляет описательный заголовок ("Запчасти"). Весь функционал выполняют дочерние методы.
        /// <br/>
        /// Работает с современной версией Excel.
        /// </summary>
        /// <param name="document">Документ, к которому принадлежит нужная страница.</param>
        /// <param name="sheet">Страница, в которую будет встроена информация.</param>
        /// <param name="request">Заказ.</param>
        private void insertRequestPartsRows_Seventh_Modern(XSSFWorkbook document, XSSFSheet sheet, Service_Request request) {
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(13, 13, 0, 1));
            XSSFCell headerCell = (XSSFCell)sheet.GetRow(13).GetCell(0);
            headerCell.CellStyle = generateCellStyle_Modern(document, true, true, 20);
            headerCell.SetCellValue("Запчасти");

            XSSFCell nameCell = (XSSFCell)sheet.GetRow(14).GetCell(0);
            nameCell.CellStyle = generateCellStyle_Modern(document, true, true, 16);
            nameCell.SetCellValue("Название запчасти:");

            XSSFCell countCell = (XSSFCell)sheet.GetRow(14).GetCell(1);
            countCell.CellStyle = generateCellStyle_Modern(document, true, true, 16);
            countCell.SetCellValue("Количество:");

            generateRequestPartsRows_Modern(document, sheet, request);
        }

        /// <summary>
        /// Формирует строку с информацией про необходимые для работы запчасти.
        /// <br/>
        /// Вызывает функцию для вставки новых строк в документ, а затем функции для вставки информации про запчасть.
        /// <br/>
        /// Работает с современной версией Excel.
        /// </summary>
        /// <param name="document">Документ, к которому принадлежит нужная страница.</param>
        /// <param name="sheet">Страница, в которую будет встроена информация.</param>
        /// <param name="request">Заказ.</param>
        private void generateRequestPartsRows_Modern(XSSFWorkbook document, XSSFSheet sheet, Service_Request request) {
            for (int i = 0; i < request.Parts_To_Request.Count; i++) { 
                // 15 в данном случае — индекс первой строки с запчастями. Он никогда не меняется.
                createPartRow_Modern(sheet, 15 + i);

                insertPartNameToCell_Modern(document, sheet, request, i);
                insertPartQuantityToCell_Modern(document, sheet, request, i);
            }
        }

        /// <summary>
        /// Создает новую строку в документе, что позволяет почти бесконечно размещать запчасти на странице отчета.
        /// <br/>
        /// Работает с современной версией Excel.
        /// </summary>
        /// <param name="sheet">Страница, в которую нужно вставить строку.</param>
        /// <param name="ind">Индекс новой строки.</param>
        private void createPartRow_Modern(XSSFSheet sheet, int ind) {
            XSSFRow row = (XSSFRow)sheet.CreateRow(ind);
            row.Height = 500;

            for (int i = 0; i < 2; i++) {
                row.CreateCell(i);
            }
        }

        /// <summary>
        /// Вставляет в ячейку данные о названии запчасти.
        /// <br/>
        /// Работает с современной версией Excel.
        /// </summary>
        /// <param name="document">Документ, к которому принадлежит нужная страница.</param>
        /// <param name="sheet">Страница, в которую будет встроена информация.</param>
        /// <param name="request">Заказ.</param>
        /// <param name="i">Индекс последней строки на текущей странице.</param>
        private void insertPartNameToCell_Modern(XSSFWorkbook document, XSSFSheet sheet, Service_Request request, int i) {
            XSSFCell nameCell = (XSSFCell)sheet.GetRow(15 + i).GetCell(0);
            nameCell.CellStyle = generateCellStyle_Modern(document, true, true, 16);
            nameCell.SetCellValue(request.Parts_To_Request.ToList()[i].Part.Part_Name);
        }

        /// <summary>
        /// Вставляет в ячейку данные о количестве запчастей.
        /// <br/>
        /// Работает с современной версией Excel.
        /// </summary>
        /// <param name="document">Документ, к которому принадлежит нужная страница.</param>
        /// <param name="sheet">Страница, в которую будет встроена информация.</param>
        /// <param name="request">Заказ.</param>
        /// <param name="i">Индекс последней строки на текущей странице.</param>
        private void insertPartQuantityToCell_Modern(XSSFWorkbook document, XSSFSheet sheet, Service_Request request, int i) {
            XSSFCell countCell = (XSSFCell)sheet.GetRow(15 + i).GetCell(1);
            countCell.CellStyle = generateCellStyle_Modern(document, true, true, 16);
            countCell.SetCellValue(request.Parts_To_Request.ToList()[i].Count);
        }
        #endregion
        #endregion

        #region Формирование стилей.

        /// <summary>
        /// Формирует стиль для ячейки документа Excel современного формата.
        /// </summary>
        /// <param name="document">Документ, к которому будет относиться стиль.</param>
        /// <param name="alignToCenter">Центрировать текст по центру по горизонтали?</param>
        /// <param name="useFont">Использовать нестандартный шрифт?</param>
        /// <param name="fontHeight">Если используется нестандартный шрифт, указывает его размер.</param>
        /// <param name="hexColor">Если используется нестандартный шрифт, hex-представление цвета будет использоваться в качестве цвета шрифта.</param>
        /// <param name="centerVertical">Центрировать текст по вертикали?</param>
        /// <returns>Сформированный по параметрам стиль.</returns>
        private XSSFCellStyle generateCellStyle_Modern(XSSFWorkbook document, bool alignToCenter, bool useFont, int fontHeight = 14, string hexColor = "#000000", bool centerVertical = true) {
            XSSFCellStyle style = (XSSFCellStyle)document.CreateCellStyle();
            style.Alignment = alignToCenter ? HorizontalAlignment.Center : HorizontalAlignment.Left;
            style.VerticalAlignment = centerVertical ? VerticalAlignment.Center : VerticalAlignment.Top;

            if (useFont) {
                XSSFFont font = (XSSFFont)document.CreateFont();
                font.FontName = fontFamily;
                font.FontHeight = fontHeight * 10;

                XSSFColor color = new XSSFColor();
                color.SetRgb(convertFromHexToRgb(hexColor));
                font.SetColor(color);

                style.SetFont(font);
            }

            return style;
        }
        #endregion
    }
}
