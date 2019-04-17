using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace ClimatLibrary
{
    public class ClimatData
    {
        public List<float> Y, Y1, Y2, DixonN, Dixon1;
        public float GrubbsN, Grubbs1, Dispersion, Asymmetry, Autocorrelation,
                     TDistribution, Fisher, FDistribution1, FDistribution2, Dispersion1, Dispersion2,
                     Student;

        public ClimatData(List<float> _Y)
        {
            Y = _Y;
            Y1 = Y.GetRange(0, Y.Count / 2);
            Y2 = Y.GetRange(Y.Count / 2 - 1, Y.Count - (Y.Count / 2));
            DixonN = new List<float>();
            Dixon1 = new List<float>();
        }
    }

    public class OutlierTest
    {
        public OutlierTest() { }

        public static ClimatData Dixon(ClimatData data)
        {
            int n = data.Y.Count - 1;

            data.DixonN.Add((data.Y[n] - data.Y[n - 1]) / (data.Y[n] - data.Y[0]));
            data.DixonN.Add((data.Y[n] - data.Y[n - 1]) / (data.Y[n] - data.Y[1]));
            data.DixonN.Add((data.Y[n] - data.Y[n - 2]) / (data.Y[n] - data.Y[1]));
            data.DixonN.Add((data.Y[n] - data.Y[n - 2]) / (data.Y[n] - data.Y[2]));
            data.DixonN.Add((data.Y[n] - data.Y[n - 2]) / (data.Y[n] - data.Y[0]));
            
            data.Dixon1.Add((data.Y[0] - data.Y[1]) / (data.Y[0] - data.Y[n]));
            data.Dixon1.Add((data.Y[0] - data.Y[1]) / (data.Y[0] - data.Y[n - 1]));
            data.Dixon1.Add((data.Y[0] - data.Y[2]) / (data.Y[0] - data.Y[n - 1]));
            data.Dixon1.Add((data.Y[0] - data.Y[2]) / (data.Y[0] - data.Y[n - 2]));
            data.Dixon1.Add((data.Y[0] - data.Y[2]) / (data.Y[0] - data.Y[n]));

            return data;
        }
        public static ClimatData Dispersion(ClimatData data)
        {
            int n = data.Y.Count;

            data.Dispersion = 0;

            for (int i = 0; i < n; i++)
            {
                data.Dispersion += (float)Math.Pow(data.Y[0] - data.Y.Average(), 2);
            }

            data.Dispersion /= n - 1;

            int n1 = data.Y1.Count;
            int n2 = data.Y2.Count;

            data.Dispersion1 = 0;
            data.Dispersion2 = 0;

            for (int i = 0; i < n1; i++)
            {
                data.Dispersion1 += (float)Math.Pow(data.Y1[0] - data.Y1.Average(), 2);
            }

            data.Dispersion1 /= n1 - 2;

            for (int i = 0; i < n2; i++)
            {
                data.Dispersion2 += (float)Math.Pow(data.Y2[0] - data.Y2.Average(), 2);
            }

            data.Dispersion2 /= n2 - 2;

            return data;
        }

        public static ClimatData Grubbs(ClimatData data)
        {
            int n = data.Y.Count - 1;

            data.GrubbsN = (data.Y[n] - data.Y.Average()) / (float)Math.Sqrt(data.Dispersion);
            data.Grubbs1 = (data.Y.Average() - data.Y[0]) / (float)Math.Sqrt(data.Dispersion);

            return data;
        }

        public static ClimatData Asymmetry(ClimatData data)
        {
            int n = data.Y.Count;

            data.Asymmetry = 0;
                    
            for (int i = 0; i < n; i++)
            {
                data.Asymmetry += (float)Math.Pow((data.Y[i] / data.Y.Average()) - 1, 3);
            }

            data.Asymmetry /= ((float)Math.Sqrt(data.Dispersion) / data.Y.Average()) * (n - 1) * (n - 2);

            return data;
        }

        public static ClimatData Autocorrelation(ClimatData data)
        {
            int n = data.Y.Count;

            float Y_Average1 = 0, Y_Average2 = 0;

            for (int i = 0; i < n - 1; i++)
            {
                Y_Average1 += data.Y[i + 1] / (n - 1);
                Y_Average2 += data.Y[i] / (n - 1);
            }

            float A = 0, B = 0, C = 0;

            for (int i = 0; i < n - 1; i++)
            {
                A += (data.Y[i] - Y_Average1) * (data.Y[i + 1] - Y_Average2);
                B += (float)Math.Pow(data.Y[i] - Y_Average1, 2);
                C += (float)Math.Pow(data.Y[i + 1] - Y_Average2, 2);
            }

            data.Autocorrelation = A / (float)Math.Sqrt(B * C);

            return data;
        }

        public static ClimatData TDistribution(ClimatData data)
        {
            int n = data.Y.Count;

            data.TDistribution = data.Autocorrelation * (float)Math.Sqrt(n - 2) / (float)Math.Sqrt(1 - (float)Math.Pow(data.Autocorrelation, 2));

            return data;
        }

        public static ClimatData Fisher(ClimatData data)
        {
            data.Fisher = data.Dispersion1 / data.Dispersion2;

            return data;
        }

        public static ClimatData FDistribution(ClimatData data)
        {
            int n1 = data.Y1.Count;
            int n2 = data.Y2.Count;

            float g;
            switch (Math.Round(data.Asymmetry * 2, MidpointRounding.AwayFromZero) / 2)
            {
                case 0.0:
                    g = 1.0f;
                    break;

                case 0.5:
                    g = 0.82f;
                    break;

                case 1.0:
                    g = 0.62f;
                    break;

                case 1.5:
                    g = 0.45f;
                    break;

                case 2.0:
                    g = 0.30f;
                    break;

                case 2.5:
                    g = 0.24f;
                    break;

                case 3.0:
                    g = 0.17f;
                    break;

                case 3.5:
                    g = 0.14f;
                    break;

                case 4.0:
                    g = 0.10f;
                    break;

                default:
                    g = 0.0f;
                    break;
            }

            data.FDistribution1 = (n1 * g)
                                  / ((1 + (2 * (float)Math.Pow(data.Asymmetry, 2)) * (1 - (float)Math.Pow(data.Asymmetry, 2)))
                                  * (1 - (1 - (float)Math.Pow(data.Asymmetry, 2 * n1)) / (n1 * (1 - (float)Math.Pow(data.Asymmetry, 2)))));
            data.FDistribution2 = (n2 * g)
                                  / ((1 + (2 * (float)Math.Pow(data.Asymmetry, 2)) * (1 - (float)Math.Pow(data.Asymmetry, 2)))
                                  * (1 - (1 - (float)Math.Pow(data.Asymmetry, 2 * n2)) / (n2 * (1 - (float)Math.Pow(data.Asymmetry, 2)))));

            return data;
        }

        public static ClimatData Student(ClimatData data)
        {
            int n1 = data.Y1.Count;
            int n2 = data.Y2.Count;

            data.Student = (data.Y1.Average() - data.Y2.Average()) / (float)Math.Sqrt(n1 * data.Dispersion1 + n2 * data.Dispersion2)
                         * (float)Math.Sqrt((n1 * n2 * (n1 + n2 - 2)) / (n1 + n2));

            return data;

            // TODO:
            // 1. Вывод как в конце методички, таблицами
            // 2. Весь процесс нужно отдельно по месяцам проводить
            // 3. Выводить по отдельным станциям
            // 4. Мб сделать чтобы несколько файлов можно было выбрать
        }
    }
}
