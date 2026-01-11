using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Spreadalonia
{
    internal class Helper
    {
        /// <summary>
        /// Run action if enabled = true.
        /// </summary>
        public static void DoIfEnabled(Action action, bool enabled)
        {
            if (!enabled) return;

            action();
        }

        /// <summary>
        /// Run async action with parameter if enabled = true.
        /// </summary>
        public static async Task DoIfEnabled<T>(Func<T, Task> action, T paramValue, bool enabled)
        {
            if (!enabled) return;

            await action(paramValue);
        }

        public static IEnumerable<T> Array2dToFlatten<T>(T[,] map)
        {
            for (int row = 0; row < map.GetLength(0); row++)
            {
                for (int col = 0; col < map.GetLength(1); col++)
                {
                    yield return map[row, col];
                }
            }
        }

        public static void DebugWriteLineMethodName([CallerMemberName] string memberName = "")
        {
            Debug.WriteLine($"{memberName}()");
        }
    }
}
