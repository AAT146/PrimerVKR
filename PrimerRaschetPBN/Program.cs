using ASTRALib;

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

            // string shablon = @"C:\Users\aat146\Documents\RastrWin3\SHABLON\динамика.rst";
            rastr.Load(RG_KOD.RG_REPL, file, shablon);

            // Объявление переменной, тип - таблицаю
            var tables = rastr.Tables;

            // Объявление объекта, содержащего таблицу "Узлы"
            var node = tables.Item("node");

            // Объявление объекта, содержащего таблицу "Generator"
            var generator = tables.Item("Generator");

            // Создание объектов, содержащих информацию по каждой колонке
            var numberNode = node.Cols.Item("ny");   // Номер узла
            var nameNode = node.Cols.Item("name");   // Название узла
            var numberArea = node.Cols.Item("na");   // Номер района
            var powerActiveLoad = node.Cols.Item("pn");   // Активная мощность нагрузки
            var powerRectiveGeneration = node.Cols.Item("qg");   // Реактивная мощность генерации
            var voltageCalc = node.Cols.Item("vras");   // Расчётное напряжение
            var deltaCalc = node.Cols.Item("delta");   // Расчётный угол

            var powerActiveGeneration = generator.Cols.Item("P");   // Активная мощность генерации.

            // Объявление объекта, содержащего таблицу "Ветви"
            var vetv = tables.Item("vetv");

            // Создание объектов, содержащих информацию по каждой колонке
            var staVetv = vetv.Cols.Item("sta");   // Состояние ветви
            var nameVetv = vetv.Cols.Item("name");   // Название ветви

            // Вывод в консоль название узлов
            for (int i = 1; i <= 46; i++)
            {
                var setSelName = "ny=" + i;   // Переменная ny = i (№ узла = i)
                node.SetSel(setSelName);   // Выборка по переменной
                var index = node.FindNextSel(-1);   // Возврат индекса след.строки, удовл-ей выборке (искл: -1)
                var nameN = nameNode.Z[index];   // Переменная с найденным индексом в столбце Название
                Console.WriteLine($"Название узла 10: {nameN}.");
            }

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

            string fileNew = @"C:\Users\aat146\Desktop\ПримерКП\Режим2.rst";
            rastr.Save(fileNew, shablon);
        }
    }
}
