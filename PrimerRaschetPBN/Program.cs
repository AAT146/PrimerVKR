// <copyright file="Program.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

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
            Rastr rst = new Rastr();

            string patch = @"C:\Users\Анастасия\Desktop\1.rst";
            rst.Load(RG_KOD.RG_REPL, patch, string.Empty);

            var tables = rst.Tables;
            var node = tables.Item("node");

            var pg = node.Cols.Item("pn");

            pg.Z[0] = 555;

            string patch_ = @"C:\Users\Анастасия\Desktop\11.rst";
            rst.Save(patch_, string.Empty);
        }
    }
}
