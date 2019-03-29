using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClimatLibrary
{
    public class OutlierTest
    {
        public OutlierTest() { }

        public static List<float> Dixon(List<float> Y)
        {
            List<float> answer = new List<float>();
            int n = Y.Count - 1;
            // Для максимального члена
            answer.Add((Y[n] - Y[n - 1]) / (Y[n] - Y[0])); // D1n
            answer.Add((Y[n] - Y[n - 1]) / (Y[n] - Y[1])); // D2n
            answer.Add((Y[n] - Y[n - 2]) / (Y[n] - Y[1])); // D3n
            answer.Add((Y[n] - Y[n - 2]) / (Y[n] - Y[2])); // D4n
            answer.Add((Y[n] - Y[n - 2]) / (Y[n] - Y[0])); // D5n

            // Для минимального члена
            answer.Add((Y[0] - Y[1]) / (Y[0] - Y[n]));     // D11
            answer.Add((Y[0] - Y[1]) / (Y[0] - Y[n - 1])); // D21
            answer.Add((Y[0] - Y[2]) / (Y[0] - Y[n - 1])); // D31
            answer.Add((Y[0] - Y[2]) / (Y[0] - Y[n - 2])); // D41
            answer.Add((Y[0] - Y[2]) / (Y[0] - Y[n]));     // D51

            return answer;
        }

        public static List<float> SmirnovGrubbs(List<float> Y)
        {
            List<float> answer = new List<float>();
            int n = Y.Count - 1;

            float Yav = Y.Average(); // среднее значение Y
            float disp = 0;          // дисперсия
            for (int i = 0; i <= n; i++)
            {
                disp += (float)Math.Pow(Y[0] - Yav, 2f) / (n - 1f);
            }

            answer.Add(Yav);
            answer.Add(disp);
            answer.Add((Y[n] - Yav) / (float)Math.Sqrt(disp)); // Gn
            answer.Add((Yav - Y[0]) / (float)Math.Sqrt(disp)); // G1

            return answer;
        }
    }
}
