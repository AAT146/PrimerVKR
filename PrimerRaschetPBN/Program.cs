using ASTRALib;
using System.Xml.Linq;

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

            // Объявление объекта, содержащего таблицу "Узлы"
            ITable Node = rastr.Tables.Item("Node");

            // Создание объектов, содержащих информацию по каждой колонке
            ICol numberNode = Node.Cols.Item("ny"); // Номер узла
            ICol nameNode = Node.Cols.Item("name"); // Название узла
            ICol numberArea = Node.Cols.Item("na"); // Номер района
            ICol powerActiveLoad = Node.Cols.Item("pn"); // Активная мощность нагрузки.
            ICol powerRectiveLoad = Node.Cols.Item("qn"); // Реактивная мощность нагрузки.
            ICol powerActiveGeneration = Node.Cols.Item("pg"); // Активная мощность генерации.
            ICol powerRectiveGeneration = Node.Cols.Item("qg"); // Реактивная мощность генерации.
            ICol voltageCalc = Node.Cols.Item("vras"); // Расчётное напряжение.
            ICol deltaCalc = Node.Cols.Item("delta"); // Расчётное напряжение.
        }
    }
}
