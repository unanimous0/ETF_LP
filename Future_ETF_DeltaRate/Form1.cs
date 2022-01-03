using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms.DataVisualization.Charting;

namespace RealtimeGraph
{
    public partial class Form1 : Form
    {
        double x;

        double FUTURE_DELTA = 0.00d;
        double ETF_DELTA    = 0.00d;
        double DELTA_DIFF   = 0.00d;

        private string filePath = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            button1.Text = "Start";
            timer1.Tick += timer1_Tick;
            timer1.Interval = 50;

            // chart1에 디폴트로 추가되어있는 Series의 차트 -> FUTURE_DELTA
            // 1. FUTURE_DELTA
            chart1.Series[0].ChartType = SeriesChartType.Line;
            chart1.Series[0].LegendText = "FUTURE_DELTA";

            // 2. ETF_DELTA
            chart1.Series.Add("ETF_DELTA");
            chart1.Series["ETF_DELTA"].ChartType = SeriesChartType.Line;

            chart1.Series.Add("DELTA_DIFF");
            chart1.ChartAreas.Add("DELTA_DIFF");
            chart1.Series["DELTA_DIFF"].ChartArea = "DELTA_DIFF";
            chart1.Series["DELTA_DIFF"].ChartType = SeriesChartType.Line;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            chart1.Series[0].Points.AddXY(x, FUTURE_DELTA);
            chart1.Series["ETF_DELTA"].Points.AddXY(x, ETF_DELTA);
            chart1.ChartAreas[0].AxisX.Minimum = chart1.Series[0].Points[0].XValue;
            chart1.ChartAreas[0].AxisX.Maximum = x;

            //if(chart1.Series[0].Points.Count > 100)
            //{
            //    chart1.Series[0].Points.RemoveAt(0);
            //}

            chart1.Series["DELTA_DIFF"].Points.AddXY(x, FUTURE_DELTA);
            chart1.ChartAreas["DELTA_DIFF"].AxisX.Minimum = chart1.Series["DELTA_DIFF"].Points[0].XValue;
            chart1.ChartAreas["DELTA_DIFF"].AxisX.Maximum = x;

            x += 0.1;

            Console.WriteLine(chart1.Series["DELTA_DIFF"].Points[0].XValue);
            Console.WriteLine(x);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (timer1.Enabled)
            {
                timer1.Stop();
                button1.Text = "Start";
            }
            else
            {
                timer1.Start();
                button1.Text = "Stop";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog OFD = new OpenFileDialog();
            if (OFD.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.Clear();
                richTextBox1.Text = OFD.FileName;
                filePath = OFD.FileName;
                Console.WriteLine("읽어온 파일 경로: " + filePath);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (filePath != "")
            {
                Excel.Application application = new Excel.Application();
                Excel.Workbook workbook = application.Workbooks.Open(Filename: @filePath);
                Excel.Worksheet worksheet_401170 = workbook.Worksheets.get_Item("401170_KBiSelect메타버스");
                //application.Visible = false;

                //double ydata1 = worksheet_401170.Range[worksheet_401170.Cells[3, 16], worksheet_401170.Cells[3, 17]];
                FUTURE_DELTA = worksheet_401170.Cells[3, 16].Value;     // 선물 변화율
                ETF_DELTA    = worksheet_401170.Cells[3, 14].Value;     // 현물 변화율
                DELTA_DIFF   = worksheet_401170.Cells[5, 13].Value;     // 선물 변화율 - 현물 변화율

                StringBuilder sb = new StringBuilder();
                sb.Append("선물 변화율  =  ");
                sb.AppendLine(string.Format("{0:f2}%", FUTURE_DELTA * 100));

                sb.Append("현물 변화율  =  ");
                sb.AppendLine(string.Format("{0:f2}%", ETF_DELTA * 100));

                sb.Append("변화율 차이  =  ");
                sb.AppendLine(string.Format("{0:f2}%", DELTA_DIFF * 100));

                sb.AppendLine();

                richTextBox2.Text = sb.ToString();

                //DeleteObject(worksheet_401170);
                //DeleteObject(workbook);
                //application.Quit();
                //DeleteObject(application);
            }
        }

        private void DeleteObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("메모리 할당을 해제하는 중 문제가 발생했습니다. " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
