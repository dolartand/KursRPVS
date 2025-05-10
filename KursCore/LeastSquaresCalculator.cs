using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace KursCore
{
    /// <summary>
    /// Реализует расчет коэффициентов линейной и полиномиальной регрессии методом наименьших квадратов.
    /// Также включает функциональность для загрузки и сохранения данных в формате INI.
    /// </summary>
    public class LeastSquaresCalculator
    {
        /// <summary>
        /// Получает или задает коллекцию X-координат точек данных.
        /// </summary>
        public List<double> DataX { get; set; }

        /// <summary>
        /// Получает или задает коллекцию Y-координат точек данных.
        /// </summary>
        public List<double> DataY { get; set; }

        /// <summary>
        /// Инициализирует новый экземпляр класса <see cref="LeastSquaresCalculator"/> с пустыми наборами данных.
        /// </summary>
        public LeastSquaresCalculator()
        {
            DataX = new List<double>();
            DataY = new List<double>();
        }

        /// <summary>
        /// Вычисляет коэффициенты 'a' и 'b' для линейной регрессии вида y = a·x + b.
        /// </summary>
        /// <param name="a">Выходной параметр: коэффициент наклона прямой (a).</param>
        /// <param name="b">Выходной параметр: свободный член уравнения (b).</param>
        /// <exception cref="InvalidOperationException">
        /// Генерируется, если:
        /// <br/>- Количество элементов в <see cref="DataX"/> и <see cref="DataY"/> не совпадает.
        /// <br/>- Количество точек данных меньше двух.
        /// <br/>- Знаменатель при вычислении коэффициентов равен нулю (например, все точки X одинаковы или лежат на вертикальной линии).
        /// </exception>
        public void Compute(out double a, out double b)
        {
            if (DataX.Count != DataY.Count)
                throw new InvalidOperationException("Количество элементов в DataX и DataY должно совпадать.");
            if (DataX.Count < 2)
                throw new InvalidOperationException("Для линейной регрессии требуется как минимум две точки данных.");

            int n = DataX.Count;
            double sumX = DataX.Sum();
            double sumY = DataY.Sum();
            double sumXX = DataX.Select(x_val => x_val * x_val).Sum(); // Изменено имя переменной для избежания конфликта
            double sumXY = DataX.Zip(DataY, (x_val, y_val) => x_val * y_val).Sum(); // Изменено имя переменной

            double denominator = n * sumXX - sumX * sumX;
            if (Math.Abs(denominator) < 1e-10)
                throw new InvalidOperationException("Знаменатель равен нулю. Невозможно вычислить коэффициенты линейной регрессии (точки могут быть коллинеарны вертикально или идентичны).");

            a = (n * sumXY - sumX * sumY) / denominator;
            b = (sumY * sumXX - sumX * sumXY) / denominator;
        }

        /// <summary>
        /// Вычисляет коэффициенты полиномиальной регрессии вида y = c[0] + c[1]*x + ... + c[degree]*x^degree.
        /// Коэффициенты возвращаются в порядке [c_0, c_1, ..., c_degree].
        /// </summary>
        /// <param name="degree">Степень полинома (m). Должна быть неотрицательной.</param>
        /// <returns>Массив коэффициентов полинома [c_0, c_1, ..., c_degree].</returns>
        /// <exception cref="InvalidOperationException">
        /// Генерируется, если:
        /// <br/>- Количество элементов в <see cref="DataX"/> и <see cref="DataY"/> не совпадает.
        /// <br/>- Количество точек данных меньше, чем (degree + 1).
        /// <br/>- Не удалось решить систему линейных уравнений (например, вырожденная матрица).
        /// </exception>
        /// <exception cref="ArgumentOutOfRangeException">Генерируется, если степень полинома отрицательная.</exception>
        public double[] ComputePolynomial(int degree)
        {
            if (DataX.Count != DataY.Count)
                throw new InvalidOperationException("Количество элементов в DataX и DataY должно совпадать.");
            if (degree < 0)
                throw new ArgumentOutOfRangeException(nameof(degree), "Степень полинома должна быть неотрицательной.");

            if (DataX.Count < degree + 1)
                throw new InvalidOperationException($"Недостаточно точек данных для полиномиальной регрессии степени {degree}. Требуется как минимум {degree + 1} точек.");

            int n_points = DataX.Count; // Переименовано для ясности
            int numCoefficients = degree + 1;

            double[,] matrixA = new double[numCoefficients, numCoefficients];
            double[] vectorB = new double[numCoefficients];

            for (int i = 0; i < numCoefficients; i++)
            {
                for (int j = 0; j < numCoefficients; j++)
                {
                    double sumXPower = 0;
                    for (int k = 0; k < n_points; k++)
                    {
                        sumXPower += Math.Pow(DataX[k], i + j);
                    }
                    matrixA[i, j] = sumXPower;
                }

                double sumYXPower = 0;
                for (int k = 0; k < n_points; k++)
                {
                    sumYXPower += DataY[k] * Math.Pow(DataX[k], i);
                }
                vectorB[i] = sumYXPower;
            }

            try
            {
                return SolveGaussianElimination(matrixA, vectorB);
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidOperationException(
                    $"Не удалось решить систему уравнений для полиномиальной регрессии (степень {degree}): {ex.Message} " +
                    "Это может произойти, если точки данных коллинеарны таким образом, что система становится вырожденной, или если степень полинома слишком высока для данных.",
                    ex);
            }
        }

        /// <summary>
        /// Решает систему линейных алгебраических уравнений Ax = B методом Гаусса-Жордана.
        /// </summary>
        /// <param name="A">Матрица коэффициентов системы A.</param>
        /// <param name="B">Вектор правых частей системы B.</param>
        /// <returns>Массив решений x.</returns>
        /// <exception cref="InvalidOperationException">
        /// Генерируется, если матрица вырожденная или система плохо обусловлена (нулевой ведущий элемент),
        /// или если вычисленный коэффициент является NaN или бесконечностью.
        /// </exception>
        private double[] SolveGaussianElimination(double[,] A, double[] B)
        {
            int n_eq = B.Length; // Переименовано для ясности
            double[,] Ab = new double[n_eq, n_eq + 1];

            for (int i = 0; i < n_eq; i++)
            {
                for (int j = 0; j < n_eq; j++) Ab[i, j] = A[i, j];
                Ab[i, n_eq] = B[i];
            }

            for (int i = 0; i < n_eq; i++)
            {
                int maxRow = i;
                for (int k = i + 1; k < n_eq; k++)
                {
                    if (Math.Abs(Ab[k, i]) > Math.Abs(Ab[maxRow, i]))
                        maxRow = k;
                }

                if (maxRow != i)
                {
                    for (int k = i; k < n_eq + 1; k++)
                    {
                        double temp = Ab[i, k];
                        Ab[i, k] = Ab[maxRow, k];
                        Ab[maxRow, k] = temp;
                    }
                }

                if (Math.Abs(Ab[i, i]) < 1e-12)
                    throw new InvalidOperationException("Вырожденная матрица или плохо обусловленная система (нулевой ведущий элемент).");

                double pivot = Ab[i, i];
                for (int k = i; k < n_eq + 1; k++)
                {
                    Ab[i, k] /= pivot;
                }

                for (int k = 0; k < n_eq; k++)
                {
                    if (k != i)
                    {
                        double factor = Ab[k, i];
                        for (int j = i; j < n_eq + 1; j++)
                        {
                            Ab[k, j] -= factor * Ab[i, j];
                        }
                    }
                }
            }

            double[] x_solution = new double[n_eq]; // Переименовано для ясности
            for (int i = 0; i < n_eq; i++)
            {
                x_solution[i] = Ab[i, n_eq];
                if (double.IsNaN(x_solution[i]) || double.IsInfinity(x_solution[i]))
                {
                    throw new InvalidOperationException($"Вычисленный коэффициент c[{i}] является NaN или Бесконечностью. Проблема с численной устойчивостью или данными. Попробуйте меньшую степень полинома.");
                }
            }
            return x_solution;
        }

        /// <summary>
        /// Загружает данные X и Y из указанного INI-файла.
        /// Ожидается, что файл содержит секцию [Data] с ключами X и Y,
        /// значения которых представляют собой числа, разделенные точкой с запятой.
        /// </summary>
        /// <param name="path">Путь к INI-файлу.</param>
        /// <exception cref="ArgumentException">Генерируется, если путь к файлу пуст или равен null.</exception>
        /// <exception cref="FileNotFoundException">Генерируется, если INI-файл не найден.</exception>
        /// <exception cref="FormatException">
        /// Генерируется, если:
        /// <br/>- Произошла ошибка при парсинге данных в INI-файле.
        /// <br/>- Количество загруженных значений X и Y не совпадает.
        /// </exception>
        public void LoadFromIni(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Путь не может быть null или пустым.", nameof(path));
            if (!File.Exists(path))
                throw new FileNotFoundException("INI файл не найден.", path);

            DataX.Clear();
            DataY.Clear();

            var lines = File.ReadAllLines(path);
            bool dataSection = false;
            string currentSection = "";

            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (trimmed.StartsWith(";") || string.IsNullOrEmpty(trimmed))
                    continue;

                if (trimmed.StartsWith("[") && trimmed.EndsWith("]"))
                {
                    currentSection = trimmed.Substring(1, trimmed.Length - 2);
                    dataSection = currentSection.Equals("Data", StringComparison.OrdinalIgnoreCase);
                    continue;
                }

                if (!dataSection)
                    continue;

                var parts = trimmed.Split(new[] { '=' }, 2);
                if (parts.Length != 2)
                    continue;

                string key = parts[0].Trim();
                string valueStr = parts[1].Trim();

                try
                {
                    var values = valueStr
                        .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(s_val => double.Parse(s_val.Trim(), CultureInfo.InvariantCulture)) // Изменено имя переменной
                        .ToList();

                    if (key.Equals("X", StringComparison.OrdinalIgnoreCase))
                        DataX = values;
                    else if (key.Equals("Y", StringComparison.OrdinalIgnoreCase))
                        DataY = values;
                }
                catch (FormatException ex)
                {
                    throw new FormatException($"Ошибка парсинга данных в INI файле для ключа '{key}': {ex.Message}. Проверьте формат (например, '1.23;4.56').", ex);
                }
            }
            if (DataX.Count != DataY.Count)
                throw new FormatException("Несоответствие количества значений X и Y, загруженных из INI файла.");
        }

        /// <summary>
        /// Сохраняет текущие данные X и Y в указанный INI-файл.
        /// Данные сохраняются в секции [Data] с ключами X и Y.
        /// </summary>
        /// <param name="path">Целевой путь для сохранения INI-файла.</param>
        /// <exception cref="ArgumentException">Генерируется, если путь к файлу пуст или равен null.</exception>
        public void SaveToIni(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Путь не может быть null или пустым.", nameof(path));

            string xLine = "X=" + string.Join(";", DataX.Select(x_val => x_val.ToString(CultureInfo.InvariantCulture)));
            string yLine = "Y=" + string.Join(";", DataY.Select(y_val => y_val.ToString(CultureInfo.InvariantCulture)));

            var lines = new List<string>
            {
                "; Сгенерировано LeastSquaresCalculator",
                "[Data]",
                xLine,
                yLine
            };

            File.WriteAllLines(path, lines);
        }
    }
}