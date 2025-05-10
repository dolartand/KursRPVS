using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace KursCore
{
    /// <summary>
    /// Реализует расчет коэффициентов линейной регрессии методом наименьших квадратов
    /// </summary>
    /// <remarks>
    /// Формулы расчета:
    /// a = (nΣxy - ΣxΣy) / (nΣx² - (Σx)²)
    /// b = (ΣyΣx² - ΣxΣxy) / (nΣx² - (Σx)²)
    /// </remarks>
    public class LeastSquaresCalculator
    {
        /// <summary>
        /// Коллекция X-координат точек данных
        /// </summary>
        /// <value>Инициализируется пустым списком при создании объекта</value>
        public List<double> DataX { get; set; }

        /// <summary>
        /// Коллекция Y-координат точек данных
        /// </summary>
        /// <value>Инициализируется пустым списком при создании объекта</value>
        public List<double> DataY { get; set; }

        /// <summary>
        /// Создает экземпляр калькулятора с пустыми наборами данных
        /// </summary>
        public LeastSquaresCalculator()
        {
            DataX = new List<double>();
            DataY = new List<double>();
        }

        /// <summary>
        /// Вычисляет коэффициенты линейной регрессии y = a·x + b
        /// </summary>
        /// <param name="a">Коэффициент наклона прямой</param>
        /// <param name="b">Свободный член уравнения</param>
        /// <exception cref="InvalidOperationException">
        /// Генерируется при:
        /// - Несовпадении количества точек X и Y
        /// - Вырожденной системе (нулевой знаменатель)
        /// </exception>
        public void Compute(out double a, out double b)
        {
            if (DataX.Count != DataY.Count)
                throw new InvalidOperationException("DataX and DataY must have the same number of elements.");

            int n = DataX.Count;
            double sumX = DataX.Sum();
            double sumY = DataY.Sum();
            double sumXX = DataX.Select(x => x * x).Sum();
            double sumXY = DataX.Zip(DataY, (x, y) => x * y).Sum();

            double denominator = n * sumXX - sumX * sumX;
            if (Math.Abs(denominator) < 1e-10)
                throw new InvalidOperationException("Denominator is zero. Cannot compute regression coefficients.");

            a = (n * sumXY - sumX * sumY) / denominator;
            b = (sumY * sumXX - sumX * sumXY) / denominator;
        }

        /// <summary>
        /// Загружает данные из INI-файла
        /// </summary>
        /// <param name="path">Путь к INI-файлу</param>
        /// <remarks>
        /// Формат файла:
        /// [Data]
        /// X=1.5;2.0;3.5 (значения через точку с запятой)
        /// Y=2.1;3.9;5.6
        /// Регистр ключей не важен. Поддерживаются комментарии через ';'
        /// </remarks>
        /// <exception cref="ArgumentException">Неверный путь</exception>
        /// <exception cref="FileNotFoundException">Файл не найден</exception>
        /// <exception cref="FormatException">Ошибка формата данных</exception>
        public void LoadFromIni(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Path cannot be null or empty.", nameof(path));

            if (!File.Exists(path))
                throw new FileNotFoundException("INI file not found.", path);

            var lines = File.ReadAllLines(path);
            bool dataSection = false;
            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (trimmed.StartsWith("[") && trimmed.EndsWith("]"))
                {
                    dataSection = trimmed.Equals("[Data]", StringComparison.OrdinalIgnoreCase);
                    continue;
                }
                if (!dataSection || trimmed.Length == 0 || trimmed.StartsWith(";"))
                    continue;

                var parts = trimmed.Split(new[] { '=' }, 2);
                if (parts.Length != 2)
                    continue;

                string key = parts[0].Trim();
                string value = parts[1].Trim();

                var values = value
                    .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => double.Parse(s.Trim(), CultureInfo.InvariantCulture))
                    .ToList();

                if (key.Equals("X", StringComparison.OrdinalIgnoreCase))
                    DataX = values;
                else if (key.Equals("Y", StringComparison.OrdinalIgnoreCase))
                    DataY = values;
            }
        }

        /// <summary>
        /// Сохраняет текущие данные в INI-файл
        /// </summary>
        /// <param name="path">Целевой путь для сохранения</param>
        /// <remarks>
        /// Формат записи:
        /// [Data]
        /// X=1.5;2.0 (значения через точку с запятой)
        /// Y=2.1;3.9
        /// Используется инвариантная культура для форматирования чисел
        /// </remarks>
        /// <exception cref="ArgumentException">Неверный путь</exception>
        /// <exception cref="UnauthorizedAccessException">Отсутствуют права на запись</exception>
        public void SaveToIni(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Path cannot be null or empty.", nameof(path));

            // Формируем строки с точкой в дробной части и разделителем ';'
            string xLine = "X=" + string.Join(";", DataX.Select(x => x.ToString(CultureInfo.InvariantCulture)));
            string yLine = "Y=" + string.Join(";", DataY.Select(y => y.ToString(CultureInfo.InvariantCulture)));

            var lines = new List<string>
            {
                "; Generated by LeastSquaresCalculator",
                "[Data]",
                xLine,
                yLine
            };

            File.WriteAllLines(path, lines);
        }
    }
}
