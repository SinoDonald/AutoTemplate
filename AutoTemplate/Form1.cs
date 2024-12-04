using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using static AutoTemplate.AutoTemplate;

namespace AutoTemplate
{
    public partial class Form1 : Form
    {
        public Form1(List<PointToMatrix> pointToMatrixs, int cols, int rows)
        {
            InitializeComponent();
            InitializeChart(pointToMatrixs, cols, rows);
        }

        private void InitializeChart(List<PointToMatrix> pointToMatrixs, int cols, int rows)
        {
            // 設置圖表參數
            ChartArea chartArea = new ChartArea();
            chartArea.AxisX.Title = "X Axis";
            chartArea.AxisY.Title = "Y Axis";
            chart1.ChartAreas.Add(chartArea);            
            chart1.ChartAreas[0].AxisY.IsReversed = true; // 翻轉Y軸
            // 設置X、Y軸的範圍
            chart1.ChartAreas[0].AxisX.Minimum = 0;
            chart1.ChartAreas[0].AxisX.Maximum = cols;
            chart1.ChartAreas[0].AxisY.Minimum = 0;
            chart1.ChartAreas[0].AxisY.Maximum = rows;

            Series series = new Series();
            series.ChartType = SeriesChartType.Point; // or SeriesChartType.Line
            series.Name = "幾何圖檢視";

            // 添加點
            foreach (PointToMatrix pointToMatrix in pointToMatrixs)
            {
                if(pointToMatrix.isRectangle == 1) { series.Points.AddXY(pointToMatrix.cols, pointToMatrix.rows); }
            }

            // 將圖顯示至Form上
            chart1.Series.Add(series);
        }
        // 取消
        private void cancelBtn_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
