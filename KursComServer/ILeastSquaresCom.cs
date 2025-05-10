using System.Runtime.InteropServices;

namespace KursComServer
{
    /// <summary>
    /// COM-интерфейс для выполнения вычислений методом наименьших квадратов.
    /// Предоставляет методы для загрузки/сохранения данных, вычисления линейной
    /// и полиномиальной регрессии.
    /// </summary>
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("39856D54-B3DE-4EC9-A547-056EFF777C35")]
    public interface ILeastSquaresCom
    {
        /// <summary>
        /// Загружает данные из указанного INI-файла.
        /// </summary>
        /// <param name="path">Путь к INI-файлу.</param>
        [DispId(1)]
        void LoadFromIni(string path);

        /// <summary>
        /// Сохраняет текущие данные в указанный INI-файл.
        /// </summary>
        /// <param name="path">Путь для сохранения INI-файла.</param>
        [DispId(2)]
        void SaveToIni(string path);

        /// <summary>
        /// Вычисляет коэффициенты линейной регрессии (y = ax + b).
        /// </summary>
        /// <param name="a">Выходной параметр: коэффициент наклона (a).</param>
        /// <param name="b">Выходной параметр: свободный член (b).</param>
        [DispId(3)]
        void Compute(out double a, out double b);

        /// <summary>
        /// Устанавливает наборы данных X и Y напрямую.
        /// </summary>
        /// <param name="xs">Массив X-координат.</param>
        /// <param name="ys">Массив Y-координат.</param>
        [DispId(4)]
        void SetData(double[] xs, double[] ys);

        /// <summary>
        /// Устанавливает степень полинома для последующих вычислений полиномиальной регрессии.
        /// </summary>
        /// <param name="degree">Степень полинома.</param>
        [DispId(5)]
        void SetPolynomialDegree(int degree);

        /// <summary>
        /// Вычисляет коэффициенты полиномиальной регрессии для указанной степени.
        /// </summary>
        /// <param name="degree">Степень полинома.</param>
        /// <returns>Массив коэффициентов полинома [c0, c1, ..., c_degree].</returns>
        [DispId(6)]
        double[] ComputePolynomialCoefficients(int degree);
    }
}