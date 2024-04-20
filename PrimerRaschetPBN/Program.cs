using ASTRALib;
using OfficeOpenXml;

namespace PrimerRaschetPBN
{
    /// <summary>
    /// Расчета ПБН на примере Бодайбинского ЭР Иркутской ОЗ.
    /// </summary>
    public class Program
    {
        /// <summary>
        /// Метод: чтение файла формата Excel.
        /// </summary>
        /// <param name="filePath">Файл Excel.</param>
        /// <returns>Массив данных.</returns>
        public static double[] ReadFileFromExcel(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var data = new double[worksheet.Dimension.Rows];

                for (int i = 1; i <= worksheet.Dimension.Rows; i++)
                {
                    data[i - 1] = worksheet.Cells[i, 1].GetValue<double>();
                }

                return data;
            }
        }

        /// <summary>
        /// Метод: расчет математического ожидания.
        /// </summary>
        /// <param name="data">Массив данны.</param>
        /// <returns>Величина МО.</returns>
        public static double MathExpectation(double[] data)
        {
            double sum = 0;
            int count = 0;

            foreach (double value in data)
            {
                sum += value;
                count++;
            }

            return sum / count;
        }

        /// <summary>
        /// Метод: расчет среднеквадратического отклонения.
        /// </summary>
        /// <param name="data">Массив данных.</param>
        /// <param name="mean">МО.</param>
        /// <returns>СКО.</returns>
        public static double StandardDeviation(double[] data, double mean)
        {
            double sumSquaredDiff = 0;
            int count = 0;

            foreach (double value in data)
            {
                sumSquaredDiff += Math.Pow(value - mean, 2);
                count++;
            }

            return Math.Sqrt(sumSquaredDiff / count);
        }

        /// <summary>
        /// Упрощенное моделирование.
        /// </summary>
        public static void Main()
        {
            // Создание указателя на экземпляр RastrWin и его запуск
            IRastr rastr = new Rastr();

            // Загрузка файл
            string file = @"C:\Users\aat146\Desktop\ПроизПрактика\Растр\Режим.rg2";
            string shablon = @"C:\Programs\RastrWin3\RastrWin3\SHABLON\режим.rg2";

            rastr.Load(RG_KOD.RG_REPL, file, shablon);

            // Объявление объекта, содержащего таблицу "Узлы"
            ITable tableNode = (ITable)rastr.Tables.Item("node");

            // Объявление объекта, содержащего таблицу "Генератор(УР)"
            ITable tableGenYR = (ITable)rastr.Tables.Item("Generator");

            // Объявление объекта, содержащего таблицу "Ветви"
            ITable tableVetv = (ITable)rastr.Tables.Item("vetv");

            // Узлы
            ICol numberNode = (ICol)tableNode.Cols.Item("ny");   // Номер
            ICol nameNode = (ICol)tableNode.Cols.Item("name");   // Название
            ICol activeGen = (ICol)tableNode.Cols.Item("pg");   // Мощность генерации
            ICol activeLoad = (ICol)tableNode.Cols.Item("pn");   // Мощность нагрузки

            // Ветви
            ICol staVetv = (ICol)tableVetv.Cols.Item("sta");   // Состояние
            ICol tipVetv = (ICol)tableVetv.Cols.Item("tip");   // Тип
            ICol nStart = (ICol)tableVetv.Cols.Item("ip");   // Номер начала
            ICol nEnd = (ICol)tableVetv.Cols.Item("iq");   // Номер конца
            ICol nParall = (ICol)tableVetv.Cols.Item("np");   // Номер параллельности
            ICol nameVetv = (ICol)tableVetv.Cols.Item("name");   // Название

            // Файл Excel генеральной совопукности
            string xlsxLoad = "C:\\Users\\aat146\\Desktop\\ПроизПрактика\\Растр\\Load.xlsx";
            string xlsxGenerator = "C:\\Users\\aat146\\Desktop\\ПроизПрактика\\Растр\\Generator.xlsx";

            // Чтение данных из файла Excel
            double[] dataLoad = ReadFileFromExcel(xlsxLoad);
            double[] dataGenerator = ReadFileFromExcel(xlsxGenerator);

            // Определение МО и СКО нагрузки
            double moLoad = MathExpectation(dataLoad);
            double skoLoad = StandardDeviation(dataLoad, moLoad);

            // Определение МО и СКО генерации
            double moGenerator = MathExpectation(dataGenerator);
            double skoGenerator = StandardDeviation(dataGenerator, moGenerator);

            // Генерация случайной величины по нормальному распределению
            Random rand = new Random();
            double a = rand.NextDouble();
            double normal = Math.Sqrt(-2.0 * Math.Log(a) * Math.Sin(2.0 * Math.PI * a));

            // Формирование случайной величины с МО и СКО
            double rdmLoad = Math.Round(moLoad + (skoLoad * normal), 2);
            double rdmGenerator = Math.Round(moGenerator + (skoGenerator * normal), 2);

            Console.WriteLine($"СВ: {normal}. СВ Нагрузка: {rdmLoad}. СВ Генерация: {rdmGenerator}");

            List<string> listNodeName = new List<string>();

            int startNode = 1;
            int endNode = 6;

            for (int i = startNode; i <= endNode; i++)
            {
                var setSelName = "ny=" + i;   // Переменная ny = i (№ узла = i)
                tableNode.SetSel(setSelName);   // Выборка по переменной
                var index = tableNode.FindNextSel[-1];   // Возврат индекса след.строки, удовл-ей выборке (искл: -1)
                var nameN = nameNode.Z[index];   // Переменная с найденным индексом в столбце Название
                listNodeName.Add(nameN);    // Добавление названия в список
            }

            // Вывод в консоль название узлов
            foreach (string i in listNodeName)
            {
                Console.WriteLine($"Лист: {listNodeName}");
            }

            //int p = 500;
            //powerActiveGeneration.Z[index] = p;

            //var setSelVetv = "ip=" + 2 + "&" + "iq=" + 3 + "&" + "np=" + 2;
            //tableVetv.SetSel(setSelVetv);
            //var number = tableVetv.FindNextSel[-1];
            //staVetv.Z[number] = 1;    // 1 - отключение; 0 -включение
            //var name1v = nameVetv.Z[number];
            //Console.WriteLine($"Название ветви: {name1v}");

            // Расчет УР
            _ = rastr.rgm("");

            // Сохранение результатов
            string fileNew = @"C:\Users\aat146\Desktop\ПроизПрактика\Растр\Режим2.rg2";
            rastr.Save(fileNew, shablon);
        }
    }
}
