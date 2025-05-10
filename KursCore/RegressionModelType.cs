using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KursCore
{
    /// <summary>
    /// Определяет тип модели регрессии.
    /// </summary>
    public enum RegressionModelType
    {
        /// <summary>
        /// Линейная регрессия (y = ax + b).
        /// </summary>
        Linear,
        /// <summary>
        /// Полиномиальная регрессия (y = c0 + c1*x + c2*x^2 + ...).
        /// </summary>
        Polynomial
    }
}
