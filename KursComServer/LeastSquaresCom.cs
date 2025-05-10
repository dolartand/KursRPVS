using KursCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace KursComServer
{
    /// <summary>
    /// COM-сервер для расчета методом наименьших квадратов
    /// </summary>
    /// <remarks>
    /// Реализует интерфейс ILeastSquaresCom через делегирование вызовов
    /// внутреннему экземпляру LeastSquaresCalculator из ядра приложения
    /// </remarks>
    [ComVisible(true)]
    [Guid("BD46D277-9C60-4843-A26C-CCDD69B25C2D")]
    [ClassInterface(ClassInterfaceType.None)]
    public class LeastSquaresCom : ILeastSquaresCom
    {
        private LeastSquaresCalculator _calc = new LeastSquaresCalculator();

        /// <summary>
        /// Загружает данные из INI-файла через внутренний калькулятор
        /// </summary>
        public void LoadFromIni(string ini) => _calc.LoadFromIni(ini);

        /// <summary>
        /// Сохраняет данные в INI-файл через внутренний калькулятор
        /// </summary>
        public void SaveToIni(string ini) => _calc.SaveToIni(ini);

        /// <summary>
        /// Вычисляет коэффициенты линейной регрессии через внутренний калькулятор
        /// </summary>
        public void Compute(out double a, out double b) => _calc.Compute(out a, out b);

        /// <summary>
        /// Устанавливает сырые данные через конвертацию массивов в списки
        /// </summary>
        /// <remarks>
        /// Выполняет проверку длины массивов перед установкой значений
        /// </remarks>
        public void SetData(double[] xs, double[] ys)
        {
            _calc.DataX = xs.ToList();
            _calc.DataY = ys.ToList();
        }
    }
}
