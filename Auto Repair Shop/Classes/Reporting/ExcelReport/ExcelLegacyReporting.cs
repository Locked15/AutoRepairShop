using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using Auto_Repair_Shop.Entities;
using Auto_Repair_Shop.Resources;

namespace Auto_Repair_Shop.Classes.Reporting {
    
    /// <summary>
    /// Класс, нужный для формирования отчёта в Excel.
    /// <br/>
    /// В данном файле представлен функционал формирования отчёта в старой версии Excel.
    /// </summary>
    public partial class ExcelReporting : KirovReporting {
        
        /// <summary>
        /// Начинает формирование отчёта в устаревшей версии Excel.
        /// </summary>
        private void generateLegacyExcelReport() {
            HSSFWorkbook document = new HSSFWorkbook();

            foreach (var request in requests) {
                HSSFSheet sheet = (HSSFSheet)document.CreateSheet(getSafeSheetName(document, request));
                insertSheetRowsAndCells_First(sheet);

                inserServiceInformationRow_Second(document, sheet, request);
                insertVehicleInformationRow_Third(document, sheet, request);
                insertVehicleImageRow_Fourth(document, sheet, request);
                insertRequestDatesRow_Fifth(document, sheet, request);
                insertRequestPartsRows_Sixth(document, sheet, request);
            }

            using (FileStream fs = new FileStream(fullFileName, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite)) {
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
        private string getSafeSheetName(HSSFWorkbook document, Service_Request request) {
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
        /// </summary>
        /// <param name="sheet">Страница документа для формирования элементов.</param>
        private void insertSheetRowsAndCells_First(HSSFSheet sheet) {
            for (int i = 0; i < 14; i++) {
                HSSFRow row = (HSSFRow)sheet.CreateRow(i);
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
        /// </summary>
        /// <param name="document">Документ, в котором находится нужная страница.</param>
        /// <param name="sheet">Страница, на которой будет размещена информация.</param>
        /// <param name="request">Заказ, содержащий информацию.</param>
        private void inserServiceInformationRow_Second(HSSFWorkbook document, HSSFSheet sheet, Service_Request request) {
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, 1));

            HSSFCell mergedCell = (HSSFCell)sheet.GetRow(0).GetCell(0);
            mergedCell.CellStyle = generateCellStyle(document, true, true, 26);
            mergedCell.SetCellValue(request.Service_Type.Name.TrimEnd('.'));
        }

        /// <summary>
        /// Формирует строку с информацией о автомобиле (бренд и название).
        /// </summary>
        /// <param name="document">Документ, к которому относится таргетированная страница.</param>
        /// <param name="sheet">Страница, в которую будет встроена информация.</param>
        /// <param name="request">Заказ, содержащий информацию.</param>
        private void insertVehicleInformationRow_Third(HSSFWorkbook document, HSSFSheet sheet, Service_Request request) {
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(1, 1, 0, 1));

            HSSFCell mergedCell = (HSSFCell)sheet.GetRow(1).GetCell(0);
            mergedCell.CellStyle = generateCellStyle(document, true, true, 20);
            mergedCell.SetCellValue($"{request.Vehicle.Name} — {request.Vehicle.Vehicle_Brand.Brand}");
        }

        #region Добавление изображения.
        private void insertVehicleImageRow_Fourth(HSSFWorkbook document, HSSFSheet sheet, Service_Request request) {
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(2, 9, 0, 1));

            // Устанавливаем изменение размера для изображения.
            HSSFPicture picture = (HSSFPicture)createPicture(document, sheet, request);
            picture.Resize(1, 1);
        }

        /// <summary>
        /// Добавляет в документ изображение и инициализирует высокоуровневые объекты для работы с ним.
        /// </summary>
        /// <param name="document">Документ, в который будет добавлено изображение.</param>
        /// <param name="sheet">Страница, на которой будет размещено изображение.</param>
        /// <param name="request">Запрос, который содержит адрес изображения.</param>
        /// <returns>Высокоуровневый объект для работы с изображением.</returns>
        private IPicture createPicture(HSSFWorkbook document, HSSFSheet sheet, Service_Request request) {
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

        #region Добавление времени.

        /// <summary>
        /// Формирует строки с датами заказа (начало и конец).
        /// </summary>
        /// <param name="document">Документ, к которому принадлежит нужная страница.</param>
        /// <param name="sheet">Страница, в которую будет встроена информация.</param>
        /// <param name="request">Заказ.</param>
        private void insertRequestDatesRow_Fifth(HSSFWorkbook document, HSSFSheet sheet, Service_Request request) {
            generateBeginDate(document, sheet, request);

            generateCompleteDate(document, sheet, request);
        }

        /// <summary>
        /// Формирует на странице строку с датой размещения заказа.
        /// </summary>
        /// <param name="document">Документ, к которому принадлежит нужная страница.</param>
        /// <param name="sheet">Страница, в которую будет встроена информация.</param>
        /// <param name="request">Заказ.</param>
        private void generateBeginDate(HSSFWorkbook document, HSSFSheet sheet, Service_Request request) {
            HSSFCell descCell = (HSSFCell)sheet.GetRow(10).GetCell(0);
            descCell.CellStyle = generateCellStyle(document, false, true, 16);
            descCell.SetCellValue($"Дата размещения:");

            HSSFCell valueCell = (HSSFCell)sheet.GetRow(10).GetCell(1);
            valueCell.CellStyle = generateCellStyle(document, false, true, 16);
            valueCell.CellStyle.Alignment = HorizontalAlignment.Right;
            valueCell.SetCellValue($"{request.Request_Date.Value:dd.MM.yyyy}.");
        }

        /// <summary>
        /// Формирует на странице строку с датой выполнения заказа.
        /// </summary>
        /// <param name="document">Документ, к которому принадлежит нужная страница.</param>
        /// <param name="sheet">Страница, в которую будет встроена информация.</param>
        /// <param name="request">Заказ.</param>
        private void generateCompleteDate(HSSFWorkbook document, HSSFSheet sheet, Service_Request request) {
            HSSFCell descCell = (HSSFCell)sheet.GetRow(11).GetCell(0);
            descCell.CellStyle = generateCellStyle(document, false, true, 16);
            descCell.SetCellValue($"Дата выполнения:");

            HSSFCell valueCell = (HSSFCell)sheet.GetRow(11).GetCell(1);
            valueCell.CellStyle = generateCellStyle(document, false, true, 16);
            valueCell.CellStyle.Alignment = HorizontalAlignment.Right;
            valueCell.SetCellValue($"{request.Request_Approx_Complete.Value:dd.MM.yyyy}.");
        }
        #endregion

        #region Добавление частей, нужных для выполнения заказа.

        /// <summary>
        /// Формирует строку с информацией про необходимые для работы запчасти.
        /// <br/>
        /// Конкретно этот метод только добавляет описательный заголовок ("Запчасти"). Весь функционал выполняют дочерние методы.
        /// </summary>
        /// <param name="document">Документ, к которому принадлежит нужная страница.</param>
        /// <param name="sheet">Страница, в которую будет встроена информация.</param>
        /// <param name="request">Заказ.</param>
        private void insertRequestPartsRows_Sixth(HSSFWorkbook document, HSSFSheet sheet, Service_Request request) {
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(12, 12, 0, 1));
            HSSFCell headerCell = (HSSFCell)sheet.GetRow(12).GetCell(0);
            headerCell.CellStyle = generateCellStyle(document, true, true, 20);
            headerCell.SetCellValue("Запчасти");
            
            HSSFCell nameCell = (HSSFCell)sheet.GetRow(13).GetCell(0);
            nameCell.CellStyle = generateCellStyle(document, true, true, 16);
            nameCell.SetCellValue("Название запчасти:");

            HSSFCell countCell = (HSSFCell)sheet.GetRow(13).GetCell(1);
            countCell.CellStyle = generateCellStyle(document, true, true, 16);
            countCell.SetCellValue("Количество:");

            generateRequestPartsRows(document, sheet, request);
        }

        /// <summary>
        /// Формирует строку с информацией про необходимые для работы запчасти.
        /// <br/>
        /// Вызывает функцию для вставки новых строк в документ, а затем функции для вставки информации про запчасть.
        /// </summary>
        /// <param name="document">Документ, к которому принадлежит нужная страница.</param>
        /// <param name="sheet">Страница, в которую будет встроена информация.</param>
        /// <param name="request">Заказ.</param>
        private void generateRequestPartsRows(HSSFWorkbook document, HSSFSheet sheet, Service_Request request) {
            for (int i = 0; i < request.Parts_To_Request.Count; i++) {
                // 13 в данном случае — индекс первой строки с запчастями. Он никогда не меняется.
                createPartRow(sheet, 13 + i);

                insertPartNameToCell(document, sheet, request, i);
                insertPartQuantityToCell(document, sheet, request, i);
            }
        }

        /// <summary>
        /// Создает новую строку в документе, что позволяет почти бесконечно размещать запчасти нв странице отчета.
        /// </summary>
        /// <param name="sheet">Страница, в которую нужно вставить строку.</param>
        /// <param name="ind">Индекс новой строки.</param>
        private void createPartRow(HSSFSheet sheet, int ind) {
            HSSFRow row = (HSSFRow)sheet.CreateRow(ind);
            row.Height = 500;

            for (int i = 0; i < 2; i++) {
                row.CreateCell(i);
            }
        }

        /// <summary>
        /// Вставляет в ячейку данные о названии запчасти.
        /// </summary>
        /// <param name="document">Документ, к которому принадлежит нужная страница.</param>
        /// <param name="sheet">Страница, в которую будет встроена информация.</param>
        /// <param name="request">Заказ.</param>
        /// <param name="i">Индекс последней строки на текущей странице.</param>
        private void insertPartNameToCell(HSSFWorkbook document, HSSFSheet sheet, Service_Request request, int i) {
            HSSFCell nameCell = (HSSFCell)sheet.GetRow(13 + i).GetCell(0);
            nameCell.CellStyle = generateCellStyle(document, true, true, 16);
            nameCell.SetCellValue(request.Parts_To_Request.ToList()[i].Part.Part_Name);
        }

        /// <summary>
        /// Вставляет в ячейку данные о количестве запчастей.
        /// </summary>
        /// <param name="document">Документ, к которому принадлежит нужная страница.</param>
        /// <param name="sheet">Страница, в которую будет встроена информация.</param>
        /// <param name="request">Заказ.</param>
        /// <param name="i">Индекс последней строки на текущей странице.</param>
        private void insertPartQuantityToCell(HSSFWorkbook document, HSSFSheet sheet, Service_Request request, int i) {
            HSSFCell countCell = (HSSFCell)sheet.GetRow(13 + i).GetCell(1);
            countCell.CellStyle = generateCellStyle(document, true, true, 16);
            countCell.SetCellValue(request.Parts_To_Request.ToList()[i].Count);
        }
        #endregion
        #endregion

        #region Формирование стилей.

        /// <summary>
        /// Формирует стиль для ячейки документа Excel старого формата.
        /// </summary>
        /// <param name="document">Документ, к которому будет прикреплен стиль.</param>
        /// <param name="alignToCenter">Выравнивать текст по центру?</param>
        /// <param name="useFont">Использовать нестандартный шрифт?</param>
        /// <param name="fontHeight">В случае использования нестандартного шрифта задает размер текста.</param>
        /// <returns>Сформированный по заданным параметрам стиль.</returns>
        private HSSFCellStyle generateCellStyle(HSSFWorkbook document, bool alignToCenter, bool useFont, int fontHeight = 14, bool centerVertical = true) {
            HSSFCellStyle style = (HSSFCellStyle)document.CreateCellStyle();
            style.Alignment = alignToCenter ? HorizontalAlignment.Center : HorizontalAlignment.Left;
            style.VerticalAlignment = centerVertical ? VerticalAlignment.Center : VerticalAlignment.Top;

            if (useFont) {
                HSSFFont font = (HSSFFont)document.CreateFont();
                font.FontName = fontFamily;
                font.FontHeight = fontHeight * 10;

                style.SetFont(font);
            }

            return style;
        }
        #endregion
    }
}
