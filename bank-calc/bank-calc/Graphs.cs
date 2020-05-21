using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ZedGraph;

namespace bank_calc
{
    public partial class Graphs: Form
    {
        private List<Payment> _payments;
        Random randomGen = new Random();
        private Color getRandomColor()
        {   
            KnownColor[] names = (KnownColor[])Enum.GetValues(typeof(KnownColor));
            KnownColor randomColorName = names[randomGen.Next(names.Length)];
            Color randomColor = Color.FromKnownColor(randomColorName);
            return randomColor;
        }
        public Graphs(List<Payment> payments)
        {
            if (payments.Count == 0) this.Close();
            InitializeComponent();
            this._payments = payments;
            drawPie();
        }

        private void drawPie()
        {
            zedGraph.GraphPane.CurveList.Clear();
            zedGraph.GraphPane.GraphObjList.Clear();

            GraphPane myPane = zedGraph.GraphPane;
            myPane.Title.Text = "Грфик платежей";
            for (int i = 0; i < _payments.Count; i++)
            {
                PieItem item = myPane.AddPieSlice(_payments[i].Dolg, getRandomColor(), 0F, _payments[i].Dolg.ToString());
                item.LabelType = PieLabelType.Value;
            }

            zedGraph.AxisChange();
            // Обновляем график
            zedGraph.Invalidate();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            drawPie();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            zedGraph.GraphPane.CurveList.Clear();
            zedGraph.GraphPane.GraphObjList.Clear();

            GraphPane myPane = zedGraph.GraphPane;

            myPane.Title.Text = "Грфик платежей";
            myPane.XAxis.Title.Text = "Выплата";
            myPane.YAxis.Title.Text = "";

            BarItem myBar = myPane.AddBar("Тело кредита", null, _payments.Select(t => t.Dolg).ToArray(), Color.Red);
            for (int i = 0; i < myBar.Points.Count; i++)
            {
                TextObj barLabel = new TextObj(myBar.Points[i].Y.ToString(), myBar.Points[i].X, myBar.Points[i].Y + 20);
                barLabel.FontSpec.Border.IsVisible = false;
                myPane.GraphObjList.Add(barLabel);
            }
            myBar.Bar.Fill = new Fill(Color.Red, Color.White, Color.Red);

            myBar = myPane.AddBar("Проценты", null, _payments.Select(t => t.Percent).ToArray(), Color.Blue);
            myBar.Bar.Fill = new Fill(Color.Blue, Color.White, Color.Blue);

            for (int i = 0; i < myBar.Points.Count; i++)
            {
                TextObj barLabel = new TextObj(myBar.Points[i].Y.ToString(), myBar.Points[i].X, myBar.Points[i].Y + 20);
                barLabel.FontSpec.Border.IsVisible = false;
                myPane.GraphObjList.Add(barLabel);
            }

            myBar.Label.IsVisible = true;

            myPane.XAxis.IsVisible = true;

            myPane.XAxis.Scale.TextLabels = _payments.Select(t=>t.Num.ToString()).ToArray();
            // Set the XAxis to Text type
            myPane.XAxis.Type = AxisType.Text;

            myPane.Chart.Fill = new Fill(Color.White, Color.FromArgb(255, 255, 166), 90F);
            myPane.Fill = new Fill(Color.FromArgb(250, 250, 255));

            zedGraph.AxisChange();
            zedGraph.Invalidate();

        }
    }
}
