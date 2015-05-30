using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjektEkonometria
{
    public partial class Form1 : Form
    {
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        OleDbDataAdapter adapter = new OleDbDataAdapter();

        public Form1()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
         
        }
        
        private void getExelY_Click(object sender, EventArgs e)
        {
            string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            path = Path.Combine(path, "daneY.ods");
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path +
                                      @";Extended Properties=""Excel 8.0;HDR=YES;""";
            OleDbConnection conn = new OleDbConnection(connectionString);
            string strCmd = "select * from [Arkusz1$A1:A28]";
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

        private void getExelX_Click(object sender, EventArgs e)
        {
            string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            path = Path.Combine(path, "daneX.ods");
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path +
                                      @";Extended Properties=""Excel 8.0;HDR=YES;""";
            OleDbConnection conn = new OleDbConnection(connectionString);
            string strCmd = "select * from [Arkusz1$A0:D28]";
            OleDbCommand cmd = new OleDbCommand(strCmd,conn);
            try
            {
                conn.Open();
                ds.Clear();
                adapter.SelectCommand = cmd;
                adapter.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
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

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[] x1 = new string[26];
            string[] x2 = new string[26];
            string[] x3 = new string[26];
            string[] x4 = new string[26];
            string[] y = new string[26];

            double[] resultY = new double[26];
            double[] resultX1 = new double[26];
            double[] resultX2 = new double[26];
            double[] resultX3 = new double[26];
            double[] resultX4 = new double[26];

            double meanY = 0;
            double meanX1 = 0;
            double meanX2 = 0;
            double meanX3 = 0;
            double meanX4 = 0;
                        
            for (int i = 0; i <26 ; i++)
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
            //textBox1.Text = meanY.ToString();

            //obliczenie wartość - różnica
            double[] valueMeanY = new double[26];
            double[] valueMeanX1 = new double[26];
            double[] valueMeanX2 = new double[26];
            double[] valueMeanX3 = new double[26];
            double[] valueMeanX4 = new double[26];

            double sumY = 0;
            double sumX1 = 0;
            double sumX2 = 0;
            double sumX3 = 0;
            double sumX4 = 0;

            //(wartosc-srednia)*(wartosc-srednia)
            double[] resultYxX1 = new double[26];
            double[] resultYxX2 = new double[26];
            double[] resultYxX3 = new double[26];
            double[] resultYxX4 = new double[26];
            //suma E
            double sumYX1 = 0;
            double sumYX2 = 0;
            double sumYX3 = 0;
            double sumYX4 = 0;
            //potega valueMeanY,valueMeanX1
            double[] valueMeanY2 = new double[26];
            double[] valueMeanX12 = new double[26];
            double[] valueMeanX22 = new double[26];
            double[] valueMeanX32 = new double[26];
            double[] valueMeanX42 = new double[26];
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
            //suma E / pomnożone pierwiastki
            double X1YdividerootmultiX1Y = 0;
            double X2YdividerootmultiX2Y = 0;
            double X3YdividerootmultiX3Y = 0;
            double X4YdividerootmultiX4Y = 0;
            //tablica korelacji dla Y
            string[] correlationArray = new string[4];
            for (int i = 0; i < 26; i++)
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
                //mnożenie y z każdym
                resultYxX1[i] = valueMeanY[i] * valueMeanX1[i];
                resultYxX2[i] = valueMeanY[i] * valueMeanX2[i];
                resultYxX3[i] = valueMeanY[i] * valueMeanX3[i];
                resultYxX4[i] = valueMeanY[i] * valueMeanX4[i];
                //teraz obliczyc sume E - suma dla mnożenia
                sumYX1 += resultYxX1[i];
                sumYX2 += resultYxX2[i];
                sumYX3 += resultYxX3[i];
                sumYX4 += resultYxX4[i];
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
                //mnożenie pierwiastków
                rootmultiX1Y = rootX1 * rootY;
                rootmultiX2Y = rootX2 * rootY;
                rootmultiX3Y = rootX3 * rootY;
                rootmultiX4Y = rootX4 * rootY;
                //suma E / pomnożone pierwiastki - współczynnik korelacji dla y,x
                X1YdividerootmultiX1Y = sumYX1 / rootmultiX1Y;
                X2YdividerootmultiX2Y = sumYX2 / rootmultiX2Y;
                X3YdividerootmultiX3Y = sumYX3 / rootmultiX3Y;
                X4YdividerootmultiX4Y = sumYX4 / rootmultiX4Y;
                
            }
            //wypełnie tablicy z danym korelacji

            correlationArray[0] = X1YdividerootmultiX1Y.ToString("0.00");
            correlationArray[1] = X2YdividerootmultiX2Y.ToString("0.00");
            correlationArray[2] = X3YdividerootmultiX3Y.ToString("0.00");
            correlationArray[3] = X4YdividerootmultiX4Y.ToString("0.00");

            for (int i = 0; i < 4; i++)
            {
                dataGridView4.Rows.Add(correlationArray[i]);
            }



        }

        
    }
}
