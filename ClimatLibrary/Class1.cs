using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ClimatLibrary
{
    public class ClimatData
    {
        public List<float> Y, DixonN, Dixon1;
        public float GrubbsN, Grubbs1, Dispersion, Asymmetry, Autocorrelation, TDistribution, Fisher;

        public ClimatData(List<float> _Y)
        {
            Y = _Y;
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
            int n = data.Y.Count;
            int m = data.Y.Count / 2;

            List<float> Y1 = data.Y.GetRange(0, m);
            List<float> Y2 = data.Y.GetRange(m - 1, data.Y.Count - m);

            float Dispersion1 = 0, Dispersion2 = 0;

            for (int i = 0; i < Y1.Count; i++)
            {
                Dispersion1 += (float)Math.Pow(Y1[0] - Y1.Average(), 2);
            }

            Dispersion1 /= Y1.Count - 2;

            for (int i = 0; i < Y2.Count; i++)
            {
                Dispersion2 += (float)Math.Pow(Y2[0] - Y2.Average(), 2);
            }

            Dispersion2 /= Y2.Count - 2;

            data.Fisher = Dispersion1 / Dispersion2;

            //MessageBox.Show("0 to " + m + "\n" + (m-1) + " to " + (data.Y.Count - m), "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            return data;
        }
    }
}
