using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace bank_calc
{
    public partial class Form1 : Form
    {
        //int k = 160;
        bool flag;
        List<Payment> payments;
        public Form1()
        {
            InitializeComponent();
            Init();
        }

        private void Init()
        {
            //dataGridView1.
            payments = new List<Payment>();
            flag = true;
            dataGridView1.ColumnCount = 8;
            dataGridView1.Columns[0].Name = "№ п.п.";
            dataGridView1.Columns[1].Name = "Дата";
            dataGridView1.Columns[2].Name = "Сумма платежа";
            dataGridView1.Columns[3].Name = "Основной долг";
            dataGridView1.Columns[4].Name = "Начисленные %";
            dataGridView1.Columns[5].Name = "Остаток";
            dataGridView1.Columns[6].Name = "Уменьшение платежа";
            dataGridView1.Columns[7].Name = "Уменьшение срока";
            currencyCB.SelectedItem = "тенге";
            dateCB.SelectedItem = "месяц";
            typeCB.SelectedItem = "Аннуитетный";
        }

        private void calc_2(int start_row)
        {
            int mon_num = payments[start_row].timeAdvance;
            int mon_left = payments.Count - 1 - mon_num;
            // if (mon_left <= start_row) return;

            double V;
            double S;
            double Q;
            double Sn;
            double Qn;
            double d;
            double p;
            double n;
            n = mon_left - start_row;
            int i;
            int j;
            
            if (typeCB.SelectedItem.ToString() == "Аннуитетный")
            {
                S = payments[start_row].Left;
                Q = float.Parse(percentTB.Text) / 100 / 12;
                V = S * Q / (1 - Math.Exp(Math.Log(1 + Q) * (-n)));

                i = 1;
                j = 0;
                Sn = S;
                while (j < n)
                {

                    Qn = Sn * Q;
                    S = V - Qn;

                    payments[i + start_row].Num = i + 1 + start_row;
                    payments[i + start_row].date = dateTimePicker1.Value.AddMonths(j);
                    payments[i + start_row].Sum = Math.Round(V, 2);
                    payments[i + start_row].Dolg = Math.Round(S, 2);
                    payments[i + start_row].Percent = Math.Round(Qn, 2);

                    Sn = Sn - S;
                    if (payments[i + start_row].Advance > 0)
                    {
                        V = V - payments[i + start_row].Advance / (n - (i + start_row) - 1);
                    }
                    payments[i + start_row].Left = Math.Round(Sn, 2);

                    i++;
                    j++;
                }
                Qn = Sn * Q;
                S = V - Qn;

                Sn = 0;

                if (flag) payments.Add(new Payment
                {
                    Num = i + 1,
                    date = dateTimePicker1.Value.AddMonths(j),
                    Sum = Math.Round(V, 2),
                    Dolg = Math.Round(S, 2),
                    Percent = Math.Round(Qn, 2),
                    Left = Math.Round(Sn, 2)
                });
                else
                {
                    payments[i + start_row].Num = i + start_row + 1;
                    payments[i + start_row].date = dateTimePicker1.Value.AddMonths(j);
                    payments[i + start_row].Sum = Math.Round(V, 2);
                    payments[i + start_row].Dolg = Math.Round(S, 2);
                    payments[i + start_row].Percent = Math.Round(Qn, 2);
                    payments[i + start_row].Left = Math.Round(Sn, 2);
                }
            }
            //Дифференцированный
            else if (typeCB.SelectedItem.ToString() == "Дифференцированный")
            {
                S = float.Parse(sumTB.Text);
                n = float.Parse(timeTB.Text);
                if (dateCB.SelectedItem.ToString() == "год") n *= 12;
                Q = float.Parse(percentTB.Text) / 100 / 12;
                d = S / n;
                i = 1;
                j = 0;
                Sn = S;

                while (j < n)
                {
                    p = Sn * Q;
                    V = d + p;

                    payments[i + start_row].Num = i + start_row + 1;
                    payments[i + start_row].date = dateTimePicker1.Value.AddMonths(j);
                    payments[i + start_row].Sum = Math.Round(V, 2);
                    payments[i + start_row].Dolg = Math.Round(d, 2);
                    payments[i + start_row].Percent = Math.Round(p, 2);

                    if (payments[i + start_row].Advance > 0)
                    {
                        d = (S - payments[i + start_row].Advance) / (n);
                    }

                    Sn = Sn - d;
                    payments[i + start_row].Left = Math.Round(Sn, 2);
                    i++;
                    j++;
                }

                p = Sn * Q;
                S = Sn - p;
                V = d + p;

                Sn = 0;

                if (flag) payments.Add(new Payment
                {
                    Num = i + 1,
                    date = dateTimePicker1.Value.AddMonths(j),
                    Sum = Math.Round(V, 2),
                    Dolg = Math.Round(d, 2),
                    Percent = Math.Round(p, 2),
                    Left = Math.Round(Sn, 2)
                });
                else
                {
                    payments[i].Num = i + 1;
                    payments[i].date = dateTimePicker1.Value.AddMonths(j);
                    payments[i].Sum = Math.Round(V, 2);
                    payments[i].Dolg = Math.Round(d, 2);
                    payments[i].Percent = Math.Round(p, 2);
                    payments[i].Left = Math.Round(Sn, 2);
                }
            }

            for (int k = payments.Count - mon_num; k < payments.Count; k++)
            {
                payments[k].Sum = 0;
                payments[k].Left = 0;
                payments[k].Percent = 0;
                payments[k].Num = 0;
                payments[k].Dolg = 0;
                payments[k].Advance = 0;
            }
            flag = true;
            //dataGridView1.DataSource = payments;
            updateGridView();
        }

        private void calculate_Credit()
        {
            double V;
            double S;
            double Q;
            double Sn;
            double Qn;
            double d;
            double p;
            double n;
            int i;
            int j;
            //int np;
            //double Sum;
            //double SumV;
            //double SumP;
            if (flag)
            {
                payments.Clear();
                n = float.Parse(timeTB.Text);
                if (dateCB.SelectedItem.ToString() == "год") n *= 12;
                for (int k = 0; k < (int)n; k++)
                {
                    payments.Add(new Payment());
                }
                payments[0].Num = 1;
                payments[0].date = dateTimePicker1.Value;
                payments[0].Left = float.Parse(sumTB.Text);
            }
            else
            {
                n = payments.Count - 1;
            }

            //Sum = 0;
            //SumV = 0;
            //SumP = 0;
            if (typeCB.SelectedItem.ToString() == "Аннуитетный")
            {
                S = float.Parse(sumTB.Text);
                Q = float.Parse(percentTB.Text) / 100 / 12;
                V = S * Q / (1 - Math.Exp(Math.Log(1 + Q) * (-n)));

                i = 1;
                j = 1;
                Sn = S;
                while (j < n)
                {

                    Qn = Sn * Q;
                    S = V - Qn;

                    payments[i].Num = i + 1;
                    payments[i].date = dateTimePicker1.Value.AddMonths(j);
                    payments[i].Sum = Math.Round(V, 2);
                    payments[i].Dolg = Math.Round(S, 2);
                    payments[i].Percent = Math.Round(Qn, 2);

                    Sn = Sn - S;
                    if (payments[i].Advance > 0)
                    {
                        V = V - payments[i].Advance / (n - i - 1);
                    }
                    payments[i].Left = Math.Round(Sn, 2);

                    i++;
                    j++;
                }
                Qn = Sn * Q;
                S = V - Qn;

                Sn = 0;

                if (flag) payments.Add(new Payment
                {
                    Num = i + 1,
                    date = dateTimePicker1.Value.AddMonths(j),
                    Sum = Math.Round(V, 2),
                    Dolg = Math.Round(S, 2),
                    Percent = Math.Round(Qn, 2),
                    Left = Math.Round(Sn, 2)
                });
                else
                {
                    payments[i].Num = i + 1;
                    payments[i].date = dateTimePicker1.Value.AddMonths(j);
                    payments[i].Sum = Math.Round(V, 2);
                    payments[i].Dolg = Math.Round(S, 2);
                    payments[i].Percent = Math.Round(Qn, 2);
                    payments[i].Left = Math.Round(Sn, 2);
                }
            }
            //Дифференцированный
            else if (typeCB.SelectedItem.ToString() == "Дифференцированный")
            {
                S = float.Parse(sumTB.Text);
                n = float.Parse(timeTB.Text);
                if (dateCB.SelectedItem.ToString() == "год") n *= 12;
                Q = float.Parse(percentTB.Text) / 100 / 12;
                d = S / n;
                i = 1;
                j = 1;
                Sn = S;

                while (j < n)
                {
                    p = Sn * Q;
                    V = d + p;

                    payments[i].Num = i + 1;
                    payments[i].date = dateTimePicker1.Value.AddMonths(j);
                    payments[i].Sum = Math.Round(V, 2);
                    payments[i].Dolg = Math.Round(d, 2);
                    payments[i].Percent = Math.Round(p, 2);

                    if (payments[i].Advance > 0)
                    {
                        d = (S - payments[i].Advance) / (n);
                    }

                    Sn = Sn - d;
                    payments[i].Left = Math.Round(Sn, 2);
                    i++;
                    j++;
                }

                p = Sn * Q;
                S = Sn - p;
                V = d + p;

                Sn = 0;

                if (flag) payments.Add(new Payment
                {
                    Num = i + 1,
                    date = dateTimePicker1.Value.AddMonths(j),
                    Sum = Math.Round(V, 2),
                    Dolg = Math.Round(d, 2),
                    Percent = Math.Round(p, 2),
                    Left = Math.Round(Sn, 2)
                });
                else
                {
                    payments[i].Num = i + 1;
                    payments[i].date = dateTimePicker1.Value.AddMonths(j);
                    payments[i].Sum = Math.Round(V, 2);
                    payments[i].Dolg = Math.Round(d, 2);
                    payments[i].Percent = Math.Round(p, 2);
                    payments[i].Left = Math.Round(Sn, 2);
                }
            }

            flag = true;
            //dataGridView1.DataSource = payments;
            updateGridView();
        }

        private void updateGridView()
        {
            dataGridView1.Rows.Clear();
            foreach (var item in payments)
            {
                string[] row = new string[] {item.Num.ToString(),
                    item.date.ToString("d"),
                    item.Sum.ToString(),
                    item.Dolg.ToString(),
                    item.Percent.ToString(),
                    item.Left.ToString(),
                    item.Advance.ToString(),
                    item.timeAdvance.ToString()};
                dataGridView1.Rows.Add(row);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                calculate_Credit();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //график
        private void button4_Click(object sender, EventArgs e)
        {
            if (payments.Count == 0) return;
            Graphs g = new Graphs(payments);
            g.Show();
        }
        
        //добавить
        private void button2_Click(object sender, EventArgs e)
        {
            double val = float.Parse(textBox4.Text);
            DateTime dt = dateTimePicker2.Value;

            for (int j = 0; j < payments.Count - 1; j++)
            {
                if (dt >= payments[j].date && dt < payments[j + 1].date)
                {
                    payments[j].Advance = val;
                    flag = false;
                    updateGridView();
                    return;
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView1.MultiSelect = true;

            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);

            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Microsoft.Office.Interop.Excel.Application();
            xlexcel.Visible = true;
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
            CR.Select();
            xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
        }
        
        

        public void Export_Data_To_Word(DataGridView DGV)
        {
            if (DGV.Rows.Count == 0) return;

                int RowCount = DGV.Rows.Count;
                int ColumnCount = DGV.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

                //add rows
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = DGV.Rows[r].Cells[c].Value;
                    }
                }

                Microsoft.Office.Interop.Word.Document wordDoc = new Microsoft.Office.Interop.Word.Document();
                wordDoc.Application.Visible = true;

                //горизонтальная ориентация страницы
                wordDoc.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape;


                dynamic oRange = wordDoc.Content.Application.Selection.Range;
                string tmp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        tmp = tmp + DataArray[r, c] + "\t";
                    }
                }

                //table format
                oRange.Text = tmp;

                object Separator = Microsoft.Office.Interop.Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();

                wordDoc.Application.Selection.Tables[1].Select();
                wordDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                wordDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                wordDoc.Application.Selection.Tables[1].Rows[1].Select();
                wordDoc.Application.Selection.InsertRowsAbove(1);
                wordDoc.Application.Selection.Tables[1].Rows[1].Select();

                //стили
                wordDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                wordDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Tahoma";
                wordDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 14;

                //добавляем заголовки для таблицы
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    wordDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = DGV.Columns[c].HeaderText;
                }

                wordDoc.Application.Selection.Tables[1].Rows[1].Select();
                wordDoc.Application.Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                foreach (Microsoft.Office.Interop.Word.Section section in wordDoc.Application.ActiveDocument.Sections)
                {
                    Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                    headerRange.Text = "Кредитный калькулятор";
                    headerRange.Font.Size = 16;
                    headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Export_Data_To_Word(dataGridView1);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            EmailForm ef = new EmailForm();
            ef.Show();
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Console.WriteLine(e.RowIndex + " " + e.ColumnIndex);
            if (e.ColumnIndex < 6)
            {
                updateGridView();
                return;
            }
            else if (e.ColumnIndex == 6)
            {
                int value = int.Parse(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
                payments[e.RowIndex].Advance = value;
                flag = false;
            }
            else if(e.ColumnIndex == 7)
            {
                int value = int.Parse(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
                payments[e.RowIndex].timeAdvance = value;
                flag = false;
                calc_2(e.RowIndex);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var total = payments.Sum(t => t.Dolg) + payments.Sum(t => t.Percent);
            string msg = "Всего зачислено:" + total;
            msg += "\nОсновной долг:" + payments.Sum(t => t.Dolg).ToString();
            msg += "\nПроцент:" + payments.Sum(t => t.Percent).ToString();
            MessageBox.Show(msg, "Информация о заёме");
        }
    }

    public class Payment
    {
        public int Num { get; set; }//номер
        public DateTime date { get; set; }//дата
        public double Sum { get; set; }//сумма платежа
        public double Dolg { get; set; }//основной долг
        public double Percent { get; set; }//начисленные %
        public double Left { get; set; }//остаток
        public double Advance { get; set; }//досрочное погашение
        public int timeAdvance { get; set; }//уменьшение срока
    }
}
