using ASTRALib;
using Excel = Microsoft.Office.Interop.Excel;

namespace PrimerRaschetPBN
{
    /// <summary>
    /// Пример расчета ПБН.
    /// </summary>
    public class Program
    {
        /// <summary>
        /// основная функция.
        /// </summary>
        public static void Main()
        {
            // Создание указателя на экземпляр RastrWin и его запуск
            Rastr rastr = new Rastr();

            // Загрузка файла с данными
            string file = @"C:\Users\aat146\Desktop\ПримерКП\Режим1.rst";
            string shablon = @"C:\Programs\RastrWin3\RastrWin3\SHABLON\динамика.rst";
            string folder = @"C:\Users\aat146\Desktop\ПримерКП";
            string fileExc = "Книга2.xlsx";
            string xlsxFile = Path.Combine(folder, fileExc);

            rastr.Load(RG_KOD.RG_REPL, file, shablon);

            // Объявление переменной, тип - таблицаю
            var tables = rastr.Tables;

            // Объявление объекта, содержащего таблицу "Узлы"
            var node = tables.Item("node");

            // Объявление объекта, содержащего таблицу "Генератор(ИД)"
            var generator = tables.Item("Generator");

            // Объявление объекта, содержащего таблицу "Генератор(ИД)"
            var area = tables.Item("area");

            // Создание объектов, содержащих информацию по каждой колонке
            var numberNode = node.Cols.Item("ny");   // Номер узла
            var nameNode = node.Cols.Item("name");   // Название узла
            var numberArea = node.Cols.Item("na");   // Номер района
            var powerActiveLoad = node.Cols.Item("pn");   // Активная мощность нагрузки
            var powerRectiveGeneration = node.Cols.Item("qg");   // Реактивная мощность генерации
            var voltageCalc = node.Cols.Item("vras");   // Расчётное напряжение
            var deltaCalc = node.Cols.Item("delta");   // Расчётный угол
            var district = node.Cols.Item("na");    // Район

            var powerActiveGeneration = generator.Cols.Item("P");   // Активная мощность генерации

            //var numberArea = area.Cols.Item("na");   // Номер района
            var nameArea = area.ColS.Item("name");   // Название района

            // Объявление объекта, содержащего таблицу "Ветви"
            var vetv = tables.Item("vetv");

            // Создание объектов, содержащих информацию по каждой колонке
            var staVetv = vetv.Cols.Item("sta");   // Состояние ветви
            var nameVetv = vetv.Cols.Item("name");   // Название ветви

            List<string> listNodeName = new List<string>();

            int startNode = 1;
            int endNode = 46;

            for (int i = startNode; i <= endNode; i++)
            {
                var setSelName = "ny=" + i;   // Переменная ny = i (№ узла = i)
                node.SetSel(setSelName);   // Выборка по переменной
                var index = node.FindNextSel(-1);   // Возврат индекса след.строки, удовл-ей выборке (искл: -1)
                var nameN = nameNode.Z[index];   // Переменная с найденным индексом в столбце Название
                listNodeName.Add(nameN);    // Добавление названия в список
            }

            // Вывод в консоль название узлов
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();

            int startArea = 1;
            int endArea = 4;
            for (int i = startArea; i <= endArea; i++)
            {
                var setSelNumber = "na=" + i;
                area.SetSel(setSelNumber);
                var index = area.FindNextSel(-1);
                var areaName = nameArea.Z[index];

                // Создаём новый лист в книге

                Excel.Worksheet worksheet = workbook.Sheets.Add();

                // Задаём название листу такое же, как ниаменование района
                worksheet.Name = areaName;

                // Получаем диапазон ячеек начиная с ячейки A1
                Excel.Range range = worksheet.Range["A1"];

                for (int j = 0; j < listNodeName.Count; j++)
                {
                    range.Offset[j, 0].Value = listNodeName[j];
                }
            }

            workbook.SaveAs(xlsxFile);
            workbook.Close();
            excelApp.Quit();

            Console.WriteLine($"Пустой файл Excel успешно создан по пути: {xlsxFile}");

            //int p = 500;
            //powerActiveGeneration.Z[index] = p;

            var setSelVetv = "ip=" + 23 + "&" + "iq=" + 1;
            vetv.SetSel(setSelVetv);
            var number = vetv.FindNextSel(-1);
            staVetv.Z[number] = 1;    // 1 - отключение; 0 -включение
            var name1v = nameVetv.Z[number];
            Console.WriteLine($"Название ветви: {name1v}.");

            // Расчет УР
            _ = rastr.rgm("");

            // Сохранение результатов
            string fileNew = @"C:\Users\aat146\Desktop\ПримерКП\Режим2.rst";
            rastr.Save(fileNew, shablon);
        }
    }
}
