using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Double;

namespace ProjektEkonometria
{
    public partial class Form1 : Form
    {
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        public Application eA;
        OleDbDataAdapter adapter = new OleDbDataAdapter();

        public Form1()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
         
        }

        private static double[,] AddLineToArray(double[,] array, int rowNumber, double[,] resultArray)
        {
            var nowaDlugosc = resultArray.GetLength(0);
            if (nowaDlugosc == 0)
            {
                nowaDlugosc = 1;
            }
            else
            {
                nowaDlugosc++;
            }

            var szerokoscTablicy = array.GetLength(1);

            var newResultArray = new double[nowaDlugosc, szerokoscTablicy];

            for (int i = 0; i < nowaDlugosc - 1; i++)
            {
                for (int j = 0; j < szerokoscTablicy; j++)
                {
                    newResultArray[i, j] = resultArray[i, j];
                }
            }

            for (int i = 0; i < szerokoscTablicy; i++)
            {
                newResultArray[nowaDlugosc - 1, i] = array[rowNumber, i];
            }

            return newResultArray;
        }
        
        //////////////////////////////////////////
        private void getExelY_Click(object sender, EventArgs e)
        {
            string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            path = Path.Combine(path, "daneY.ods");
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path +
                                      @";Extended Properties=""Excel 8.0;HDR=YES;""";
            OleDbConnection conn = new OleDbConnection(connectionString);
            string strCmd = "select * from [Arkusz1$A0:A28]";
            OleDbCommand cmd = new OleDbCommand(strCmd, conn);
            try
            {
                conn.Open();
                ds2.Clear();
                adapter.SelectCommand = cmd;
                adapter.Fill(ds2);
                dataGridView2.DataSource = ds2.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
            finally
            {
                conn.Close();
            }

        }
        public virtual MathNet.Numerics.LinearAlgebra.Double.Matrix Inverse();

        private void getExelX_Click(object sender, EventArgs e)
        {
            string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            path = Path.Combine(path, "daneX.ods");
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path +
                                      @";Extended Properties=""Excel 8.0;HDR=YES;""";
            OleDbConnection conn2 = new OleDbConnection(connectionString);
            string strCmd2 = "select * from [Arkusz1$A0:D28]";
            OleDbCommand cmd2 = new OleDbCommand(strCmd2,conn2);
            try
            {
                conn2.Open();
                ds.Clear();
                adapter.SelectCommand = cmd2;
                adapter.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
            finally
            {
                conn2.Close();
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[] x1 = new string[27];
            string[] x2 = new string[27];
            string[] x3 = new string[27];
            string[] x4 = new string[27];
            string[] y = new string[27];

            double[] resultY = new double[27];
            double[] resultX1 = new double[27];
            double[] resultX2 = new double[27];
            double[] resultX3 = new double[27];
            double[] resultX4 = new double[27];

            double meanY = 0;
            double meanX1 = 0;
            double meanX2 = 0;
            double meanX3 = 0;
            double meanX4 = 0;
                        
            for (int i = 0; i <27 ; i++)
            {
                x1[i] = dataGridView1.Rows[i].Cells[0].Value.ToString();
                x2[i] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                x3[i] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                x4[i] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                y[i] = dataGridView2.Rows[i].Cells[0].Value.ToString();
                //SUMA Y + konwersja do double
                resultY[i] = Convert.ToDouble(y[i]);
                meanY = resultY[i] + meanY;
                //SUMA X1
                resultX1[i] = Convert.ToDouble(x1[i]);
                meanX1 = resultX1[i] + meanX1;
                //SYMA X2
                resultX2[i] = Convert.ToDouble(x2[i]);
                meanX2 = resultX2[i] + meanX2;
                //SUMA X3
                resultX3[i] = Convert.ToDouble(x3[i]);
                meanX3 = resultX3[i] + meanX3;
                //SUMA X4
                resultX4[i] = Convert.ToDouble(x4[i]);
                meanX4 = resultX4[i] + meanX4;

                                
            }
            meanY = meanY / resultY.Length;
            meanX1 = meanX1 / resultX1.Length;
            meanX2 = meanX2 / resultX2.Length;
            meanX3 = meanX3 / resultX3.Length;
            meanX4 = meanX4 / resultX4.Length;

            //obliczenie wartość - różnica
            double[] valueMeanY = new double[27];
            double[] valueMeanX1 = new double[27];
            double[] valueMeanX2 = new double[27];
            double[] valueMeanX3 = new double[27];
            double[] valueMeanX4 = new double[27];

            double sumY = 0;
            double sumX1 = 0;
            double sumX2 = 0;
            double sumX3 = 0;
            double sumX4 = 0;

            //(wartosc-srednia)*(wartosc-srednia)
            double[] resultYxX1 = new double[27];
            double[] resultYxX2 = new double[27];
            double[] resultYxX3 = new double[27];
            double[] resultYxX4 = new double[27];
            //X1:X
            double[] resultX1xX1 = new double[27];
            double[] resultX1xX2 = new double[27];
            double[] resultX1xX3 = new double[27];
            double[] resultX1xX4 = new double[27];
            //X2:X
            double[] resultX2xX1 = new double[27];
            double[] resultX2xX2 = new double[27];
            double[] resultX2xX3 = new double[27];
            double[] resultX2xX4 = new double[27];
            //X3:X
            double[] resultX3xX1 = new double[27];
            double[] resultX3xX2 = new double[27];
            double[] resultX3xX3 = new double[27];
            double[] resultX3xX4 = new double[27];
            //X4:X
            double[] resultX4xX1 = new double[27];
            double[] resultX4xX2 = new double[27];
            double[] resultX4xX3 = new double[27];
            double[] resultX4xX4 = new double[27];
            ///////////////////////////
            //suma Y x X1:4
            double sumYX1 = 0;
            double sumYX2 = 0;
            double sumYX3 = 0;
            double sumYX4 = 0;
            //suma x1 x X1:4
            double sumX1X1 = 0;
            double sumX1X2 = 0;
            double sumX1X3 = 0;
            double sumX1X4 = 0;
            //suma x2 x X2:4
            double sumX2X1 = 0;
            double sumX2X2 = 0;
            double sumX2X3 = 0;
            double sumX2X4 = 0;
            //suma x3 x X3:4
            double sumX3X1 = 0;
            double sumX3X2 = 0;
            double sumX3X3 = 0;
            double sumX3X4 = 0;
            //suma x4 x X4:4
            double sumX4X1 = 0;
            double sumX4X2 = 0;
            double sumX4X3 = 0;
            double sumX4X4 = 0;
            //potega valueMeanY,valueMeanX1
            double[] valueMeanY2 = new double[27];
            double[] valueMeanX12 = new double[27];
            double[] valueMeanX22 = new double[27];
            double[] valueMeanX32 = new double[27];
            double[] valueMeanX42 = new double[27];
            //suma wartosci potegi valueMeanY,summValueMeanX12
            double summValueMeanY2 = 0;
            double summValueMeanX12 = 0;
            double summValueMeanX22 = 0;
            double summValueMeanX32 = 0;
            double summValueMeanX42 = 0;
            //pierwiastek
            double rootY = 0;
            double rootX1 = 0;
            double rootX2 = 0;
            double rootX3 = 0;
            double rootX4 = 0;
            //mnozenie pierwiastków
            double rootmultiX1Y = 0;
            double rootmultiX2Y = 0;
            double rootmultiX3Y = 0;
            double rootmultiX4Y = 0;

            double rootmultiX1X1 = 0;
            double rootmultiX2X1 = 0;
            double rootmultiX3X1 = 0;
            double rootmultiX4X1 = 0;

            double rootmultiX1X2 = 0;
            double rootmultiX2X2 = 0;
            double rootmultiX3X2 = 0;
            double rootmultiX4X2 = 0;

            double rootmultiX1X3 = 0;
            double rootmultiX2X3 = 0;
            double rootmultiX3X3 = 0;
            double rootmultiX4X3 = 0;

            double rootmultiX1X4 = 0;
            double rootmultiX2X4 = 0;
            double rootmultiX3X4 = 0;
            double rootmultiX4X4 = 0;
            //suma E / pomnożone pierwiastki
            double X1YdividerootmultiX1Y = 0;
            double X2YdividerootmultiX2Y = 0;
            double X3YdividerootmultiX3Y = 0;
            double X4YdividerootmultiX4Y = 0;
            //suma E / pomnożone pierwiastki
            double X1X1dividerootmultiX1X1 = 0;
            double X2X1dividerootmultiX2X1 = 0;
            double X3X1dividerootmultiX3X1 = 0;
            double X4X1dividerootmultiX4X1 = 0;
            //suma E / pomnożone pierwiastki
            double X1X2dividerootmultiX1X2 = 0;
            double X2X2dividerootmultiX2X2 = 0;
            double X3X2dividerootmultiX3X2 = 0;
            double X4X2dividerootmultiX4X2 = 0;
            //suma E / pomnożone pierwiastki
            double X1X3dividerootmultiX1X3 = 0;
            double X2X3dividerootmultiX2X3 = 0;
            double X3X3dividerootmultiX3X3 = 0;
            double X4X3dividerootmultiX4X3 = 0;
            //suma E / pomnożone pierwiastki
            double X1X4dividerootmultiX1X4 = 0;
            double X2X4dividerootmultiX2X4 = 0;
            double X3X4dividerootmultiX3X4 = 0;
            double X4X4dividerootmultiX4X4 = 0;
            //tablica korelacji dla Y
            string[] correlationArrayY = new string[4];
            string[] correlationArrayX1 = new string[4];
            string[] correlationArrayX2 = new string[4];
            string[] correlationArrayX3 = new string[4];
            string[] correlationArrayX4 = new string[4];

            for (int i = 0; i < 27; i++)
            {
                //wartość- średnia dla wszystich
                valueMeanY[i] = resultY[i] - meanY;
                valueMeanX1[i] = resultX1[i] - meanX1;
                valueMeanX2[i] = resultX2[i] - meanX2;
                valueMeanX3[i] = resultX3[i] - meanX3;
                valueMeanX4[i] = resultX4[i] - meanX4;

                sumY = valueMeanY[i] + sumY;
                sumX1 = valueMeanX1[i] + sumX1;
                sumX2 = valueMeanX2[i] + sumX2;
                sumX3 = valueMeanX3[i] + sumX3;
                sumX4 = valueMeanX4[i] + sumX4;
                //mnożenie y z każdym y:x1/x2/x3/x4
                resultYxX1[i] = valueMeanY[i] * valueMeanX1[i];
                resultYxX2[i] = valueMeanY[i] * valueMeanX2[i];
                resultYxX3[i] = valueMeanY[i] * valueMeanX3[i];
                resultYxX4[i] = valueMeanY[i] * valueMeanX4[i];
                //mnożenie y z każdym X1:x1/x2/x3/x4
                resultX1xX1[i] = valueMeanX1[i] * valueMeanX1[i];
                resultX1xX2[i] = valueMeanX1[i] * valueMeanX2[i];
                resultX1xX3[i] = valueMeanX1[i] * valueMeanX3[i];
                resultX1xX4[i] = valueMeanX1[i] * valueMeanX4[i];
                //mnożenie y z każdym X2:x1/x2/x3/x4
                resultX2xX1[i] = valueMeanX2[i] * valueMeanX1[i];
                resultX2xX2[i] = valueMeanX2[i] * valueMeanX2[i];
                resultX2xX3[i] = valueMeanX2[i] * valueMeanX3[i];
                resultX2xX4[i] = valueMeanX2[i] * valueMeanX4[i];
                //mnożenie y z każdym X3:x1/x2/x3/x4
                resultX3xX1[i] = valueMeanX3[i] * valueMeanX1[i];
                resultX3xX2[i] = valueMeanX3[i] * valueMeanX2[i];
                resultX3xX3[i] = valueMeanX3[i] * valueMeanX3[i];
                resultX3xX4[i] = valueMeanX3[i] * valueMeanX4[i];
                //mnożenie y z każdym X4:x1/x2/x3/x4
                resultX4xX1[i] = valueMeanX4[i] * valueMeanX1[i];
                resultX4xX2[i] = valueMeanX4[i] * valueMeanX2[i];
                resultX4xX3[i] = valueMeanX4[i] * valueMeanX3[i];
                resultX4xX4[i] = valueMeanX4[i] * valueMeanX4[i];

                //teraz obliczyc sume WYNIKU MNOZENIA - suma dla mnożenia
                sumYX1 += resultYxX1[i];
                sumYX2 += resultYxX2[i];
                sumYX3 += resultYxX3[i];
                sumYX4 += resultYxX4[i];

                sumX1X1 += resultX1xX1[i];
                sumX1X2 += resultX1xX2[i];
                sumX1X3 += resultX1xX3[i];
                sumX1X4 += resultX1xX4[i];

                sumX2X1 += resultX2xX1[i];
                sumX2X2 += resultX2xX2[i];
                sumX2X3 += resultX2xX3[i];
                sumX2X4 += resultX2xX4[i];

                sumX3X1 += resultX3xX1[i];
                sumX3X2 += resultX3xX2[i];
                sumX3X3 += resultX3xX3[i];
                sumX3X4 += resultX3xX4[i];

                sumX4X1 += resultX4xX1[i];
                sumX4X2 += resultX4xX2[i];
                sumX4X3 += resultX4xX3[i];
                sumX4X4 += resultX4xX4[i];

                //potega C i D oddzielnie
                valueMeanY2[i] = valueMeanY[i] * valueMeanY[i];
                valueMeanX12[i] = valueMeanX1[i] * valueMeanX1[i];
                valueMeanX22[i] = valueMeanX2[i] * valueMeanX2[i];
                valueMeanX32[i] = valueMeanX3[i] * valueMeanX3[i];
                valueMeanX42[i] = valueMeanX4[i] * valueMeanX4[i];
                //suma poteg C i D oddzielnie
                summValueMeanY2 += valueMeanY2[i];
                summValueMeanX12 += valueMeanX12[i];
                summValueMeanX22 += valueMeanX22[i];
                summValueMeanX32 += valueMeanX32[i];
                summValueMeanX42 += valueMeanX42[i];
                //pierwiastek sumy poteg C i D oddzielnie
                rootY = Math.Sqrt(summValueMeanY2);
                rootX1 = Math.Sqrt(summValueMeanX12);
                rootX2 = Math.Sqrt(summValueMeanX22);
                rootX3 = Math.Sqrt(summValueMeanX32);
                rootX4 = Math.Sqrt(summValueMeanX42);
                //mnożenie pierwiastków DLA Y;X1:4
                rootmultiX1Y = rootX1 * rootY;
                rootmultiX2Y = rootX2 * rootY;
                rootmultiX3Y = rootX3 * rootY;
                rootmultiX4Y = rootX4 * rootY;
                //mnożenie pierwiastków DLA X1:4;X1:4
                rootmultiX1X1 = rootX1 * rootX1;
                rootmultiX2X1 = rootX2 * rootX1;
                rootmultiX3X1 = rootX3 * rootX1;
                rootmultiX4X1 = rootX4 * rootX1;

                rootmultiX1X2 = rootX1 * rootX2;
                rootmultiX2X2 = rootX2 * rootX2;
                rootmultiX3X2 = rootX3 * rootX2;
                rootmultiX4X2 = rootX4 * rootX2;

                rootmultiX1X3 = rootX1 * rootX3;
                rootmultiX2X3 = rootX2 * rootX3;
                rootmultiX3X3 = rootX3 * rootX3;
                rootmultiX4X3 = rootX4 * rootX3;

                rootmultiX1X4 = rootX1 * rootX4;
                rootmultiX2X4 = rootX2 * rootX4;
                rootmultiX3X4 = rootX3 * rootX4;
                rootmultiX4X4 = rootX4 * rootX4;
                //suma E / pomnożone pierwiastki - współczynnik korelacji dla y,x
                X1YdividerootmultiX1Y = sumYX1 / rootmultiX1Y;
                X2YdividerootmultiX2Y = sumYX2 / rootmultiX2Y;
                X3YdividerootmultiX3Y = sumYX3 / rootmultiX3Y;
                X4YdividerootmultiX4Y = sumYX4 / rootmultiX4Y;
                //suma E / pomnożone pierwiastki - współczynnik korelacji dla y,x
                X1X1dividerootmultiX1X1 = sumX1X1 / rootmultiX1X1;
                X2X1dividerootmultiX2X1 = sumX1X2 / rootmultiX2X1;
                X3X1dividerootmultiX3X1 = sumX1X3 / rootmultiX3X1;
                X4X1dividerootmultiX4X1 = sumX1X4 / rootmultiX4X1;
                //suma E / pomnożone pierwiastki - współczynnik korelacji dla y,x
                X1X2dividerootmultiX1X2 = sumX2X1 / rootmultiX1X2;
                X2X2dividerootmultiX2X2 = sumX2X2 / rootmultiX2X2;
                X3X2dividerootmultiX3X2 = sumX2X3 / rootmultiX3X2;
                X4X2dividerootmultiX4X2 = sumX2X4 / rootmultiX4X2;
                //suma E / pomnożone pierwiastki - współczynnik korelacji dla y,x
                X1X3dividerootmultiX1X3 = sumX3X1 / rootmultiX1X3;
                X2X3dividerootmultiX2X3 = sumX3X2 / rootmultiX2X3;
                X3X3dividerootmultiX3X3 = sumX3X3 / rootmultiX3X3;
                X4X3dividerootmultiX4X3 = sumX3X4 / rootmultiX4X3;
                //suma E / pomnożone pierwiastki - współczynnik korelacji dla y,x
                X1X4dividerootmultiX1X4 = sumX4X1 / rootmultiX1X4;
                X2X4dividerootmultiX2X4 = sumX4X2 / rootmultiX2X4;
                X3X4dividerootmultiX3X4 = sumX4X3 / rootmultiX3X4;
                X4X4dividerootmultiX4X4 = sumX4X4 / rootmultiX4X4;
                                
            }
            //wypełnie tablicy z danym korelacji

            correlationArrayY[0] = X1YdividerootmultiX1Y.ToString("0.000");
            correlationArrayY[1] = X2YdividerootmultiX2Y.ToString("0.000");
            correlationArrayY[2] = X3YdividerootmultiX3Y.ToString("0.000");
            correlationArrayY[3] = X4YdividerootmultiX4Y.ToString("0.000");

            correlationArrayX1[0] = X1X1dividerootmultiX1X1.ToString("0.000");
            correlationArrayX1[1] = X2X1dividerootmultiX2X1.ToString("0.000");
            correlationArrayX1[2] = X3X1dividerootmultiX3X1.ToString("0.000");
            correlationArrayX1[3] = X4X1dividerootmultiX4X1.ToString("0.000");

            correlationArrayX2[0] = X1X2dividerootmultiX1X2.ToString("0.000");
            correlationArrayX2[1] = X2X2dividerootmultiX2X2.ToString("0.000");
            correlationArrayX2[2] = X3X2dividerootmultiX3X2.ToString("0.000");
            correlationArrayX2[3] = X4X2dividerootmultiX4X2.ToString("0.000");

            correlationArrayX3[0] = X1X3dividerootmultiX1X3.ToString("0.000");
            correlationArrayX3[1] = X2X3dividerootmultiX2X3.ToString("0.000");
            correlationArrayX3[2] = X3X3dividerootmultiX3X3.ToString("0.000");
            correlationArrayX3[3] = X4X3dividerootmultiX4X3.ToString("0.000");

            correlationArrayX4[0] = X1X4dividerootmultiX1X4.ToString("0.000");
            correlationArrayX4[1] = X2X4dividerootmultiX2X4.ToString("0.000");
            correlationArrayX4[2] = X3X4dividerootmultiX3X4.ToString("0.000");
            correlationArrayX4[3] = X4X4dividerootmultiX4X4.ToString("0.000");

                        
            for (int i = 0; i < 4; i++)
            {
                dataGridView4.Rows.Add(correlationArrayY[i]);
                dataGridView3.Rows.Add(correlationArrayX1[i], correlationArrayX2[i], correlationArrayX3[i], correlationArrayX4[i]);
                
            }

            string[,] r = new string[4,4];

            for (int i = 0; i < 4; i++)
            {
                r[i, 0] = correlationArrayX1[i];
                r[i, 1] = correlationArrayX2[i];
                r[i, 2] = correlationArrayX3[i];
                r[i, 3] = correlationArrayX4[i]; 
            }

            //konwersja tablicy wielowymiarowej R0 do double
            double[,] rD = new double[4,4];
            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < 4; j++)
                {
                    rD[i, j] = Convert.ToDouble(r[i, j]); 
                }

            }

            //poszukiwnie najwyzszej wartosci dla wiersza X1
            double[] highestInRow = new double[4];
            double rt = 1;

            highestInRow[0] = rD[0, 0];
            for (int i = 0; i < 4; i++){
                if (highestInRow[0] >= rt)
                {
                    highestInRow[0] = rD[i, 0];
                }
            }

            for (int i = 0; i < 3; i++)
            {
                if (highestInRow[0] < rD[i + 1, 0])
                {
                    highestInRow[0] = rD[i + 1, 0];
                }
            }

            //poszukiwnie najwyzszej wartosci dla wiersza X2
            highestInRow[1] = rD[0, 1];

            for (int i = 0; i < 3; i++)
            {
                if (highestInRow[1] < rD[i + 1, 1])
                {
                    if (rD[i+1,1] >= rt)
                    {
                        i = i + 1;
                    }
                    highestInRow[1] = rD[i + 1, 1];
                }
            }

            //poszukiwnie najwyzszej wartosci dla wiersza X3
            highestInRow[2] = rD[0, 2];

            for (int i = 0; i < 3; i++)
            {
                if (highestInRow[2] < rD[i + 1, 2])
                {
                    if (rD[i + 1, 2] >= rt)
                    {
                        i = i + 1;
                    }
                    highestInRow[2] = rD[i + 1, 2];
                }
            }

            //poszukiwnie najwyzszej wartosci dla wiersza X4
            highestInRow[3] = rD[0, 3];
            for (int i = 0; i < 2; i++)
            {
                if (highestInRow[3] < rD[i + 1, 3])
                {
                    highestInRow[3] = rD[i + 1, 3];
                }
            }
            //poszukiwnie najwyzszej wartosci dla kolumny highestInRow

            double highestInCol = highestInRow[0];
            for (int i = 0; i < 3; i++)
            {
                if (highestInCol > highestInRow[i + 1])
                {
                    highestInCol = highestInRow[i + 1];
                }
            }
            string highestInColSting;
            highestInColSting = Convert.ToString(highestInCol);
            rStar.Text = highestInColSting;


            //zerowanie wartosci dla rD mniejszych lub równych od rStar(wartości krytycznej)
            //highestInCol
            //rD
            double absRD = 0;
            double[,] zeroRD = new double[4,4];
            absRD = Math.Abs(highestInCol);

            zeroRD = rD;
            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < 4; j++)
                {
                    double test=0;
                    test = zeroRD[i,j];
                    test = Math.Abs(test);

                    if (test <= absRD)
                    {
                        zeroRD[i, j] = 0;
                    }
                }

            }

            //sprawdzenie który z x należy do modelu
            int[] count = new int[4];
            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < 4; j++)
                {
                    if (zeroRD[i, j] != 0)
                    {
                        count[i] += 1;
                    }
                }

            }

            //Transponowanie macierzy X'
            double[,] resultTab = new double[27, 27];
            double[,] resultTran = new double[27, 27];
            
            for (int i = 0; i < 27; i++)//tworzenie tablicy wielowymiarowej
            {
                int k = 0;
                if (count[k] > count[k + 1])
                {
                    resultTab[i, 0] = resultX1[i];
                }
                k += 1; 

                if (count[k] > count[k + 1])
                {
                 resultTab[i, 1] = resultX2[i];
                }
                k += 1;

                if (count[k] >= count[k + 1])
                {
                    resultTab[i, 2] = resultX3[i];
                    resultTab[i, 3] = resultX4[i];
                }         

            }
            //////////////////////////////////////////////

            double[,] oneTab = new double[1, 27];
            for (int i = 0; i < 27; i++)
            {
                oneTab[0, i] = 1;
            }

            //dodanie do tablicy x wchodzących do modelu tablicy jedynek
            var summOneArray = new double[27, 27];
            for (int i = 0; i < 27; i++)
            {
                for (int j = 0; j < 27; j++)
                {
                    summOneArray[i, j] = resultTab[i, j];
                }
            }
            for (int i = 0; i < 27; i++)
            {
                summOneArray[i, 4] = oneTab[0, i];
            }

            //transponuj
            for (int i = 0; i < 27; i++)
            {
                for (int j = 0; j < 27; j++)
                {
                    resultTran[j, i] = summOneArray[i, j];                    
                }
            }

            //kasowanie zer z transponowanej
            var array = resultTran;

            var dlugoscTablicy = array.GetLength(0);
            var szerokoscTablicy = array.GetLength(1);

            var resultArray = new double[0, szerokoscTablicy];

            bool isOnlyZero = true;

            for (int i = 0; i < dlugoscTablicy; i++)
            {
                for (int j = 0; j < szerokoscTablicy; j++)
                {
                    if (array[i, j] != 0)
                    {
                        isOnlyZero = false;
                        break;
                    }
                }

                if (!isOnlyZero)
                {
                    resultArray = AddLineToArray(array, i, resultArray);
                }

                isOnlyZero = true;
            }
            //transponuj do normalnej

            double[,] testtest = new double[27,3];
            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 27; j++)
                {
                    testtest[j, i] = resultArray[i, j];
                    
                }
            }

            //transponowana * y
            double[,] transRazyY = new double[27, 27];
            double[,] y2w = new double[1, 27];

            for (int i = 0; i < 27; i++)
            {
                y2w[0, i] = resultY[i];
            }

            for (int i = 0; i < 27; i++)
            {
                for (int k = 0; k < 27; k++)
                {
                    for (int j = 0; j < 3; j++)
                    {
                        transRazyY[i, k] += y2w[0, k] * testtest[k,j];//wynik macierzy transponowanej * y macierz 27*27
                    }
                }
            }
            //trans*nie trans 
            double[,] m1m2 = new double[27, 27];
            for (int i = 0; i < 27; i++)
            {
                for (int k = 0; k < 27; k++)
                {
                    for (int j = 0; j < 3; j++)
                    {
                        m1m2[i, k] += testtest[i, j] * resultArray[j, k];//wynik maceirz transponowana * nie transponowana
                    }
                }
            }

            //odwracanie macierzy m1m2

            double[,] onesMatrix = new double[27, 27];
            for (int i = 0; i < 1; i++)
            {                
                for (int j = 0; j < 27; j++)
                {
                    onesMatrix[i,j] = 1;
                    i = i + 1;
                }
            }
        }
        
    }
}
