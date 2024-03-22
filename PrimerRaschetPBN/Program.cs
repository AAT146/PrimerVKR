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
            string patch = @"C:\Users\aat146\Desktop\ПримерКП\Режим1.rst";
            rastr.Load(RG_KOD.RG_REPL, patch, " ");

            // Объявление переменной, тип - таблицаю
            var tables = rastr.Tables;

            // Объявление объекта, содержащего таблицу "Узлы"
            var node = tables.Item("Node");

            // Создание объектов, содержащих информацию по каждой колонке
            var numberNode = node.Cols.Item("ny"); // Номер узла
            var nameNode = node.Cols.Item("name"); // Название узла
            var numberArea = node.Cols.Item("na"); // Номер района
            var powerActiveLoad = node.Cols.Item("pn"); // Активная мощность нагрузки.
            var powerRectiveLoad = node.Cols.Item("qn"); // Реактивная мощность нагрузки.
            var powerActiveGeneration = node.Cols.Item("pg"); // Активная мощность генерации.
            var powerRectiveGeneration = node.Cols.Item("qg"); // Реактивная мощность генерации.
            var voltageCalc = node.Cols.Item("vras"); // Расчётное напряжение.
            var deltaCalc = node.Cols.Item("delta"); // Расчётный угол.

            // Объявление объекта, содержащего таблицу "Ветви"
            var vetv = tables.Item("Vetv");

            // Создание объектов, содержащих информацию по каждой колонке
            var staVetv = vetv.Cols.Item("sta"); // Состояние ветви.

            // Изменение Рг в строку 2, таблицы Узлы
            var setSelVoltage = "ny=" + 10;
            node.SetSel(setSelVoltage);
            var nodeNumber = node.FindNextSel(-1);
            var u10 = voltageCalc.Z[nodeNumber];
            Console.WriteLine($"Напряжение в узле 10 равно: {u10} кВ.");

            // Расчет УР
            _ = rastr.rgm(" ");

            string patchNew = @"C:\Users\aat146\Desktop\ПримерКП\Режим2.rst";
            rastr.Save(patchNew, " ");
        }
    }
}
