using ASTRALib;
using System;
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
            string fileExc = "Результат.xlsx";
            string xlsxFile = Path.Combine(folder, fileExc);

            rastr.Load(RG_KOD.RG_REPL, file, shablon);

            // Вывод в консоль название узлов
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();

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

            var areaNumber = area.Cols.Item("na");   // Номер района
            var nameArea = area.ColS.Item("name");   // Название района

            // Объявление объекта, содержащего таблицу "Ветви"
            var vetv = tables.Item("vetv");

            // Создание объектов, содержащих информацию по каждой колонке
            var staVetv = vetv.Cols.Item("sta");   // Состояние ветви
            var nameVetv = vetv.Cols.Item("name");   // Название ветви

            // Лист с названиями районов
            List<string> listAreaName = new List<string>();

            // Лист с порядковыми номерами районов
            List<int> listAreaNumber = new List<int>();

            // Цикл заполнения листов
            int startArea = 0;
            int endArea = 3;
            int startNode = 0;
            int endNode = 45;

            for (int i = startArea; i <= endArea; i++)
            {
                var nameA = nameArea.Z[i];
                var numberA = areaNumber.Z[i];
                listAreaNumber.Add(numberA);
                listAreaName.Add(nameA);
            }

            for (int i = startArea; i <= endArea; i++)
            {
                // Создаём новый лист в книге
                Excel.Worksheet sheet0 = workbook.Sheets.Add();
                sheet0.Name = listAreaName[0];
                Excel.Worksheet sheet1 = workbook.Sheets.Add();
                sheet1.Name = listAreaName[1];
                Excel.Worksheet sheet2 = workbook.Sheets.Add();
                sheet2.Name = listAreaName[2];
                Excel.Worksheet sheet3 = workbook.Sheets.Add();
                sheet3.Name = listAreaName[3];

                if (numberArea.Z[i] == listAreaNumber[0])
                {
                    // Получаем диапазон ячеек начиная с ячейки A1
                    Excel.Range range = sheet0.Range["A1"];

                    range.Offset[0, 0].Value = "Название узла";
                    range.Offset[0, 1].Value = "Номер района";
                    range.Offset[0, 2].Value = "Название района";

                    for (int j = startNode; j <= endNode; j++)
                    {
                        //range.Offset[j + 1, 0].Value = nameNode[j];
                        range.Offset[j + 1, 1].Value = listAreaNumber[0];
                        range.Offset[j + 1, 2].Value = listAreaName[0];
                    }
                }

                if (numberArea.Z[i] == listAreaNumber[1])
                {
                    // Получаем диапазон ячеек начиная с ячейки A1
                    Excel.Range range = sheet1.Range["A1"];

                    range.Offset[0, 0].Value = "Название узла";
                    range.Offset[0, 1].Value = "Номер района";
                    range.Offset[0, 2].Value = "Название района";

                    for (int j = startNode; j <= endNode; j++)
                    {
                        //range.Offset[j + 1, 0].Value = nameNode[j];
                        range.Offset[j + 1, 1].Value = listAreaNumber[1];
                        range.Offset[j + 1, 2].Value = listAreaName[1];
                    }
                }

                if (numberArea.Z[i] == listAreaNumber[2])
                {
                    // Получаем диапазон ячеек начиная с ячейки A1
                    Excel.Range range = sheet2.Range["A1"];

                    range.Offset[0, 0].Value = "Название узла";
                    range.Offset[0, 1].Value = "Номер района";
                    range.Offset[0, 2].Value = "Название района";

                    for (int j = startNode; j <= endNode; j++)
                    {
                        //range.Offset[j + 1, 0].Value = nameNode[j];
                        range.Offset[j + 1, 1].Value = listAreaNumber[2];
                        range.Offset[j + 1, 2].Value = listAreaName[2];
                    }
                }

                if (numberArea.Z[i] == listAreaNumber[3])
                {
                    // Получаем диапазон ячеек начиная с ячейки A1
                    Excel.Range range = sheet3.Range["A1"];

                    range.Offset[0, 0].Value = "Название узла";
                    range.Offset[0, 1].Value = "Номер района";
                    range.Offset[0, 2].Value = "Название района";

                    for (int j = startNode; j <= endNode; j++)
                    {
                        //range.Offset[j + 1, 0].Value = nameNode[j];
                        range.Offset[j + 1, 1].Value = listAreaNumber[3];
                        range.Offset[j + 1, 2].Value = listAreaName[3];
                    }
                }
            }

            //// Заполнение листа
            //int startNode = 0;
            //int endNode = 45;
            //for (int index = startNode; index <= endNode; index++)
            //{
            //    //var setSelName = "ny=" + i;
            //    //node.SetSel(setSelName);
            //    //var index = node.FindNextSel(-1);

            //    if (numberArea.Z[index] == listAreaNumber[0])
            //    {
            //        // Создаём новый лист в книге
            //        Excel.Worksheet worksheet = workbook.Sheets.Add();

            //        // Задаём название листу такое же, как ниаменование района
            //        //worksheet.Name = listAreaName[0];

            //        // Получаем диапазон ячеек начиная с ячейки A1
            //        Excel.Range range = worksheet.Range["A1"];

            //        range.Offset[0, 0].Value = "Название узла";
            //        range.Offset[0, 1].Value = "Номер района";
            //        range.Offset[0, 2].Value = "Название района";

            //        for (int j = 0; j <= endNode; j++)
            //        {
            //            //range.Offset[j + 1, 0].Value = nameNode[index];
            //            range.Offset[j + 1, 1].Value = listAreaNumber[0];
            //            range.Offset[j + 1, 2].Value = listAreaName[0];
            //        }
            //    }

            //    if (numberArea.Z[index] == listAreaNumber[1])
            //    {
            //        // Создаём новый лист в книге
            //        Excel.Worksheet worksheet = workbook.Sheets.Add();

            //        // Задаём название листу такое же, как ниаменование района
            //        //worksheet.Name = listAreaName[1];

            //        // Получаем диапазон ячеек начиная с ячейки A1
            //        Excel.Range range = worksheet.Range["A1"];

            //        range.Offset[0, 0].Value = "Название узла";
            //        range.Offset[0, 1].Value = "Номер района";
            //        range.Offset[0, 2].Value = "Название района";

            //        for (int j = 0; j <= endNode; j++)
            //        {
            //            //range.Offset[j + 1, 0].Value = nameNode[j];
            //            range.Offset[j + 1, 1].Value = listAreaNumber[1];
            //            range.Offset[j + 1, 2].Value = listAreaName[1];
            //        }
            //    }

            //    if (numberArea.Z[index] == listAreaNumber[2])
            //    {
            //        // Создаём новый лист в книге
            //        Excel.Worksheet worksheet = workbook.Sheets.Add();

            //        // Задаём название листу такое же, как ниаменование района
            //        //worksheet.Name = listAreaName[2];

            //        // Получаем диапазон ячеек начиная с ячейки A1
            //        Excel.Range range = worksheet.Range["A1"];

            //        range.Offset[0, 0].Value = "Название узла";
            //        range.Offset[0, 1].Value = "Номер района";
            //        range.Offset[0, 2].Value = "Название района";

            //        for (int j = 0; j <= endNode; j++)
            //        {
            //            //range.Offset[j + 1, 0].Value = nameNode[index];
            //            range.Offset[j + 1, 1].Value = listAreaNumber[2];
            //            range.Offset[j + 1, 2].Value = listAreaName[2];
            //        }
            //    }

            //    if (numberArea.Z[index] == listAreaNumber[3])
            //    {
            //        // Создаём новый лист в книге
            //        Excel.Worksheet worksheet = workbook.Sheets.Add();

            //        // Задаём название листу такое же, как ниаменование района
            //        //worksheet.Name = listAreaName[3];

            //        // Получаем диапазон ячеек начиная с ячейки A1
            //        Excel.Range range = worksheet.Range["A1"];

            //        range.Offset[0, 0].Value = "Название узла";
            //        range.Offset[0, 1].Value = "Номер района";
            //        range.Offset[0, 2].Value = "Название района";

            //        for (int j = 0; j <= endNode; j++)
            //        {
            //            //range.Offset[j + 1, 0].Value = nameNode[index];
            //            range.Offset[j + 1, 1].Value = listAreaNumber[3];
            //            range.Offset[j + 1, 2].Value = listAreaName[3];
            //        }
            //    }
            //}

            workbook.SaveAs(xlsxFile);
            workbook.Close();
            excelApp.Quit();

            Console.WriteLine($"Файл Excel успешно создан по пути: {xlsxFile}");

            // Расчет УР
            _ = rastr.rgm("");

            // Сохранение результатов
            string fileNew = @"C:\Users\aat146\Desktop\ПримерКП\Режим2.rst";
            rastr.Save(fileNew, shablon);
        }
    }
}
