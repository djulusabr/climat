using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
                //data.Asymmetry += (float)Math.Pow(data.Y[i] - data.Y.Average(), 3);
            }

            //data.Asymmetry /= n * (float)Math.Pow((float)Math.Sqrt(data.Dispersion), 3);

            data.Asymmetry *= n;

            data.Asymmetry /= (float)Math.Pow((float)Math.Sqrt(data.Dispersion) / data.Y.Average(), 3) * (n - 1) * (n - 2);

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

            float closest = Table2_C.Aggregate((x, y) => Math.Abs(x - data.Asymmetry) < Math.Abs(y - data.Asymmetry) ? x : y);
            float g = 0;
            for (int i = 0; i < Table2_C.Count; i++)
                if (Table2_C[i] == closest)
                    g = Table2_G[i];

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
            // 1) Вывод как в конце методички, таблицами
            // 2) Вывод в эксель
            // а) Мб сделать чтобы несколько файлов можно было выбрать
        }

        private static readonly ReadOnlyCollection<float> Table2_C =
            new ReadOnlyCollection<float>(new[]
                {
                    0f,
                    0.5f,
                    1.0f,
                    1.5f,
                    2.0f,
                    2.5f,
                    3.0f,
                    3.5f,
                    4.0f
                });

        private static readonly ReadOnlyCollection<float> Table2_G =
            new ReadOnlyCollection<float>(new[]
                {
                    1.0f,
                    0.82f,
                    0.62f,
                    0.45f,
                    0.30f,
                    0.24f,
                    0.17f,
                    0.14f,
                    0.10f
            });

        private static readonly ReadOnlyCollection<int> Table1_N =
            new ReadOnlyCollection<int>(new[]
                {
                    10,
                    11,
                    12,
                    13,
                    14,
                    15,
                    16,
                    17,
                    18,
                    19,
                    20,
                    21,
                    22,
                    23,
                    24,
                    25,
                    26,
                    27,
                    28,
                    29,
                    30,
                    35,
                    40,
                    50,
                    60,
                    70,
                    80,
                    90,
                    100,
                    120,
                    150,
                    200,
                    250,
                    300
            });

        private static readonly ReadOnlyCollection<float> Table1_A5 =
            new ReadOnlyCollection<float>(new[]
                {
                    0.576f,
                    0.553f,
                    0.532f,
                    0.514f,
                    0.497f,
                    0.482f,
                    0.468f,
                    0.456f,
                    0.444f,
                    0.433f,
                    0.423f,
                    0.413f,
                    0.404f,
                    0.396f,
                    0.388f,
                    0.381f,
                    0.374f,
                    0.367f,
                    0.361f,
                    0.355f,
                    0.349f,
                    0.325f,
                    0.304f,
                    0.275f,
                    0.250f,
                    0.232f,
                    0.217f,
                    0.205f,
                    0.195f,
                    0.178f,
                    0.159f,
                    0.138f,
                    0.124f,
                    0.113f
            });

        private static readonly ReadOnlyCollection<float> Table1_A1 =
            new ReadOnlyCollection<float>(new[]
                {
                    0.708f,
                    0.684f,
                    0.661f,
                    0.641f,
                    0.623f,
                    0.606f,
                    0.590f,
                    0.575f,
                    0.561f,
                    0.549f,
                    0.537f,
                    0.526f,
                    0.515f,
                    0.505f,
                    0.496f,
                    0.487f,
                    0.478f,
                    0.470f,
                    0.463f,
                    0.456f,
                    0.449f,
                    0.418f,
                    0.393f,
                    0.354f,
                    0.325f,
                    0.302f,
                    0.283f,
                    0.267f,
                    0.254f,
                    0.232f,
                    0.208f,
                    0.181f,
                    0.162f,
                    0.148f
            });

        private static readonly ReadOnlyCollection<float> Table3_R =
            new ReadOnlyCollection<float>(new[]
                {
                    0.00f,
                    0.01f,
                    0.02f,
                    0.03f,
                    0.04f,
                    0.05f,
                    0.06f,
                    0.07f,
                    0.08f,
                    0.09f,
                    0.10f,
                    0.11f,
                    0.12f,
                    0.13f,
                    0.14f,
                    0.15f,
                    0.16f,
                    0.17f,
                    0.18f,
                    0.19f,
                    0.20f,
                    0.21f,
                    0.22f,
                    0.23f,
                    0.24f,
                    0.25f,
                    0.26f,
                    0.27f,
                    0.28f,
                    0.29f,
                    0.30f,
                    0.31f,
                    0.32f,
                    0.33f,
                    0.34f,
                    0.35f,
                    0.36f,
                    0.37f,
                    0.38f,
                    0.39f,
                    0.40f,
                    0.41f,
                    0.42f,
                    0.43f,
                    0.44f,
                    0.45f,
                    0.46f,
                    0.47f,
                    0.48f,
                    0.49f,
                    0.50f,
                    0.51f,
                    0.52f,
                    0.53f,
                    0.54f,
                    0.55f,
                    0.56f,
                    0.57f,
                    0.58f,
                    0.59f,
                    0.60f,
                    0.61f,
                    0.62f,
                    0.63f,
                    0.64f,
                    0.65f,
                    0.66f,
                    0.67f,
                    0.68f,
                    0.69f
            });

        private static readonly ReadOnlyCollection<float> Table3_C =
            new ReadOnlyCollection<float>(new[]
                {
                    1.00f,
                    1.00f,
                    1.01f,
                    1.02f,
                    1.02f,
                    1.03f,
                    1.04f,
                    1.05f,
                    1.06f,
                    1.07f,
                    1.07f,
                    1.08f,
                    1.09f,
                    1.10f,
                    1.11f,
                    1.12f,
                    1.13f,
                    1.14f,
                    1.16f,
                    1.17f,
                    1.18f,
                    1.19f,
                    1.20f,
                    1.22f,
                    1.23f,
                    1.24f,
                    1.26f,
                    1.27f,
                    1.29f,
                    1.31f,
                    1.33f,
                    1.34f,
                    1.35f,
                    1.37f,
                    1.39f,
                    1.41f,
                    1.42f,
                    1.44f,
                    1.46f,
                    1.49f,
                    1.51f,
                    1.52f,
                    1.54f,
                    1.56f,
                    1.59f,
                    1.67f,
                    1.63f,
                    1.65f,
                    1.68f,
                    1.70f,
                    1.72f,
                    1.75f,
                    1.78f,
                    1.81f,
                    1.84f,
                    1.88f,
                    1.92f,
                    1.95f,
                    1.99f,
                    2.03f,
                    2.06f,
                    2.07f,
                    2.13f,
                    2.17f,
                    2.21f,
                    2.24f,
                    2.28f,
                    2.32f,
                    2.36f,
                    2.40f
            });
    }
}