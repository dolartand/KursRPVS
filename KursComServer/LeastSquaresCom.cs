using KursCore;
using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace KursComServer
{
    /// <summary>
    /// COM-сервер, реализующий интерфейс <see cref="ILeastSquaresCom"/> для выполнения
    /// вычислений методом наименьших квадратов.
    /// Является оберткой над классом <see cref="KursCore.LeastSquaresCalculator"/>.
    /// </summary>
    [ComVisible(true)]
    [Guid("BD46D277-9C60-4843-A26C-CCDD69B25C2D")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("KursComServer.LeastSquaresCom")]
    public class LeastSquaresCom : ILeastSquaresCom
    {
        private readonly LeastSquaresCalculator _calc = new LeastSquaresCalculator();
        private int _polynomialDegree = 2;

        /// <summary>
        /// Загружает данные из указанного INI-файла, используя внутренний калькулятор.
        /// </summary>
        /// <param name="path">Путь к INI-файлу.</param>
        public void LoadFromIni(string path)
        {
            _calc.LoadFromIni(path);
        }

        /// <summary>
        /// Сохраняет текущие данные в указанный INI-файл, используя внутренний калькулятор.
        /// </summary>
        /// <param name="path">Путь для сохранения INI-файла.</param>
        public void SaveToIni(string path)
        {
            _calc.SaveToIni(path);
        }

        /// <summary>
        /// Вычисляет коэффициенты линейной регрессии (y = ax + b), используя внутренний калькулятор.
        /// </summary>
        /// <param name="a">Выходной параметр: коэффициент наклона (a).</param>
        /// <param name="b">Выходной параметр: свободный член (b).</param>
        public void Compute(out double a, out double b)
        {
            _calc.Compute(out a, out b);
        }

        /// <summary>
        /// Устанавливает наборы данных X и Y напрямую во внутренний калькулятор.
        /// </summary>
        /// <param name="xs">Массив X-координат.</param>
        /// <param name="ys">Массив Y-координат.</param>
        /// <exception cref="ArgumentNullException">Выбрасывается, если один из входных массивов равен null.</exception>
        /// <exception cref="ArgumentException">Выбрасывается, если длины массивов X и Y не совпадают.</exception>
        public void SetData(double[] xs, double[] ys)
        {
            if (xs == null || ys == null)
                throw new ArgumentNullException(nameof(xs) + " или " + nameof(ys), "Входные массивы не могут быть null."); // Уточнено сообщение
            if (xs.Length != ys.Length)
                throw new ArgumentException("Массивы X и Y должны иметь одинаковую длину.");

            _calc.DataX = xs.ToList();
            _calc.DataY = ys.ToList();
        }

        /// <summary>
        /// Устанавливает степень полинома для использования при вызове <see cref="ComputePolynomialCoefficients(int)"/>,
        /// если этот метод будет вызван без явного указания степени, или для внутреннего использования (в данной реализации не используется сохраненное значение).
        /// </summary>
        /// <param name="degree">Степень полинома. Должна быть неотрицательной.</param>
        /// <exception cref="ArgumentOutOfRangeException">Выбрасывается, если степень полинома отрицательная.</exception>
        public void SetPolynomialDegree(int degree)
        {
            if (degree < 0)
                throw new ArgumentOutOfRangeException(nameof(degree), "Степень полинома должна быть неотрицательной.");
            _polynomialDegree = degree;
        }

        /// <summary>
        /// Вычисляет коэффициенты полиномиальной регрессии для указанной степени, используя внутренний калькулятор.
        /// </summary>
        /// <param name="degree">Степень полинома. Должна быть неотрицательной.</param>
        /// <returns>Массив коэффициентов полинома [c0, c1, ..., c_degree].</returns>
        /// <exception cref="ArgumentOutOfRangeException">Выбрасывается, если степень полинома отрицательная.</exception>
        public double[] ComputePolynomialCoefficients(int degree)
        {
            if (degree < 0)
                throw new ArgumentOutOfRangeException(nameof(degree), "Степень полинома для COM вызова должна быть неотрицательной.");
            return _calc.ComputePolynomial(degree);
        }
    }
}