using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.CSharp.RuntimeBinder;
using System.Windows; 
using System.Threading;

namespace Alfredas
{
    public partial class Form1 : Form
    {
        int List1CheckedItemIndex;
        Boolean List1HasChanged;
        Boolean FinishedLoading;
        Panel forBox;
        DataGridView Archyvas;
        List<WorkerInfo> Arch_workers; 
        List<WorkerInfo> workers; 
        Boolean CanExit;
        public Form1()
        {
            FinishedLoading = false;
            List1HasChanged = false;
            List1CheckedItemIndex = -2;
            InitializeComponent();
            FillListBoxWorkers();
            List1CheckedItemIndex = -1;
            FinishedLoading = true;
            CanExit = true;
            LanguageNustatymas();
        }
        public void FillListBoxWorkers()
        {
            Arch_workers = new List<WorkerInfo>();
            workers = new List<WorkerInfo>();
            Archyvas = new DataGridView();
            Archyvas.AllowUserToResizeRows = false;
            Archyvas.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            Archyvas.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            Archyvas.Dock = DockStyle.Fill;
            Archyvas.Location = new System.Drawing.Point(4, 4);
            Archyvas.ReadOnly = true;
            for (int i = 1; i < dataGridView1.Columns.Count-1; i++)
            {
                Archyvas.Columns.Add(dataGridView1.Columns[i].Name, dataGridView1.Columns[i].HeaderText);
            }
            Archyvas.Columns.Add("Data", "Data");
            Archyvas.Columns.Add(dataGridView1.Columns[dataGridView1.Columns.Count - 1].Name, dataGridView1.Columns[dataGridView1.Columns.Count - 1].HeaderText);
            string location = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\Duom.txt";
            location = location.Substring(6);
            System.IO.StreamReader file = null;
            Boolean Empty = false;
            try
            {
               file = new System.IO.StreamReader(location, false);
            }
            catch
            {
                WorkerInfo wi = new WorkerInfo();
                wi.mxkd = 0; wi.Rodikliai = new List<List<string>>() { }; wi.Uzduotys = new List<List<string>>() { };
                Empty = true;
                dataGridView1.Rows.Add(false, "Marko", 1000.0, 0, 0, 1000.0, "( ? )"); workers.Add(wi);
                
                wi = new WorkerInfo(); wi.mxkd = 0; wi.Rodikliai = new List<List<string>>() { }; wi.Uzduotys = new List<List<string>>() { };
                dataGridView1.Rows.Add(false, "Employee 2", 2000.0, 0, 0, 2000.0, "( ? )"); workers.Add(wi);
               
                wi = new WorkerInfo(); wi.mxkd = 0; wi.Rodikliai = new List<List<string>>() { }; wi.Uzduotys = new List<List<string>>() { };
                dataGridView1.Rows.Add(false, "Employee 3", 3000.0, 0, 0, 3000.0, "( ? )"); workers.Add(wi);


                dataGridView2.Rows.Add(false, "PROFIT", 100000.0, 130000.0, 121000.0, 1200.0);
                dataGridView2.Rows.Add(false, "COSTS", 500000.0, 450000.0, 480000.0, 800.0);
                dataGridView2.Rows.Add(false, "Indicator 3", 2.0, 4.0, 2.4, 2.0);
                dataGridView2.Rows.Add(false, "Indicator 4", 3.0, 6.0, 3.6, 3.0);

                dataGridView3.Rows.Add(false, "RELATIONSHIP WITH...", 11.0, 9.0);
                dataGridView3.Rows.Add(false, "CUTTING ON THE RET...", 9.0, 5.0);
                dataGridView3.Rows.Add(false, "IMPROVING PROFESS...", 7.0, 6.0);
                dataGridView3.Rows.Add(false, "TASK 4", 8.0, 7.0);
                dataGridView3.Rows.Add(false, "TASK 5", 9.0, 8.0);

                wi = new WorkerInfo(); wi.mxkd = 0; wi.Rodikliai = new List<List<string>>() { }; wi.Uzduotys = new List<List<string>>() { };
                Archyvas.Rows.Add("Marko", 1000.00, 0, 0, 1000.00, "2014.06.16", "( ? )"); Arch_workers.Add(wi);

                wi = new WorkerInfo(); wi.mxkd = 0; wi.Rodikliai = new List<List<string>>() { }; wi.Uzduotys = new List<List<string>>() { };
                Archyvas.Rows.Add("Employee 2", 2000.00, 0, 0, 2000.00, "2014.06.16", "( ? )"); Arch_workers.Add(wi);

                wi = new WorkerInfo(); wi.mxkd = 0; wi.Rodikliai = new List<List<string>>() { }; wi.Uzduotys = new List<List<string>>() { };
                Archyvas.Rows.Add("Employee 3", 3000.00, 0, 0, 3000.00, "2014.06.16", "( ? )"); Arch_workers.Add(wi);
            }
            if (Empty == false)
            {
                bool GalimaKonvertuoti = TikrintiArReikiaKonvertuoti();
                if (GalimaKonvertuoti == false)
                {
                    while (file.EndOfStream != true)
                    {
                        string eilute = file.ReadLine();
                        if (eilute == "Darbuotoju lentele:")
                        {
                            int index = Convert.ToInt32(file.ReadLine());
                            for (int i = 0; i < index; i++)
                            {
                                string[] reiksmes = file.ReadLine().Split(',');
                                dataGridView1.Rows.Add(false, reiksmes[0], Convert.ToDouble(reiksmes[1]), Convert.ToDouble(reiksmes[2]), Convert.ToDouble(reiksmes[3]), Convert.ToDouble(reiksmes[4]), "( ? )");
                                reiksmes = file.ReadLine().Split(',');
                                int rod = Convert.ToInt32(reiksmes[0]);
                                int uzd = Convert.ToInt32(reiksmes[1]);
                                WorkerInfo wi = new WorkerInfo();
                                wi.mxkd = Convert.ToDouble(reiksmes[2]);
                                for (int j = 0; j < rod; j++)
                                {
                                    reiksmes = file.ReadLine().Split(',');
                                    wi.Rodikliai.Add(new List<string>(reiksmes));
                                }
                                for (int j = 0; j < uzd; j++)
                                {
                                    reiksmes = file.ReadLine().Split(',');
                                    wi.Uzduotys.Add(new List<string>(reiksmes));
                                }
                                workers.Add(wi);
                            }
                        }
                        else if (eilute == "Rodikliu lentele:")
                        {
                            int index = Convert.ToInt32(file.ReadLine());
                            for (int i = 0; i < index; i++)
                            {
                                string[] reiksmes = file.ReadLine().Split(',');
                                dataGridView2.Rows.Add(false, reiksmes[0], Convert.ToDouble(reiksmes[1]), Convert.ToDouble(reiksmes[2]), Convert.ToDouble(reiksmes[3]), Convert.ToDouble(reiksmes[4])); ;
                            }
                        }
                        else if (eilute == "Uzduociu lentele:")
                        {
                            int index = Convert.ToInt32(file.ReadLine());
                            for (int i = 0; i < index; i++)
                            {
                                string[] reiksmes = file.ReadLine().Split(',');
                                dataGridView3.Rows.Add(false, reiksmes[0], Convert.ToDouble(reiksmes[1]), Convert.ToDouble(reiksmes[2]));
                            }
                        }
                        else if (eilute == "Archyvo lentele:")
                        {
                            int index = Convert.ToInt32(file.ReadLine());
                            for (int i = 0; i < index; i++)
                            {
                                string eiluteSK = file.ReadLine();
                                string[] reiksmes = eiluteSK.Split(',');
                                Archyvas.Rows.Add(reiksmes[0], Convert.ToDouble(reiksmes[1]), Convert.ToDouble(reiksmes[2]), Convert.ToDouble(reiksmes[3]), Convert.ToDouble(reiksmes[4]), reiksmes[5], "( ? )");
                                reiksmes = file.ReadLine().Split(',');
                                int rod = Convert.ToInt32(reiksmes[0]);
                                int uzd = Convert.ToInt32(reiksmes[1]);
                                WorkerInfo wi = new WorkerInfo();
                                wi.mxkd = Convert.ToDouble(reiksmes[2]);
                                for (int j = 0; j < rod; j++)
                                {
                                    reiksmes = file.ReadLine().Split(',');
                                    wi.Rodikliai.Add(new List<string>(reiksmes));
                                }
                                for (int j = 0; j < uzd; j++)
                                {
                                    reiksmes = file.ReadLine().Split(',');
                                    wi.Uzduotys.Add(new List<string>(reiksmes));
                                }
                                Arch_workers.Add(wi);
                            }
                        }
                    }
                }
                else
                {
                    while (file.EndOfStream != true)
                    {
                        string eilute = file.ReadLine();
                        if (eilute == "Darbuotoju lentele:")
                        {
                            int index = Convert.ToInt32(file.ReadLine());
                            for (int i = 0; i < index; i++)
                            {
                                string[] reiksmes = file.ReadLine().Split(',');
                                dataGridView1.Rows.Add(false, Konvertuoti(reiksmes[0], ".", ","), Convert.ToDouble(Konvertuoti(reiksmes[1], ".", ",")), Convert.ToDouble(Konvertuoti(reiksmes[2], ".", ",")), Convert.ToDouble(Konvertuoti(reiksmes[3], ".", ",")), Convert.ToDouble(Konvertuoti(reiksmes[4], ".", ",")), "( ? )");
                                reiksmes = file.ReadLine().Split(',');
                                int rod = Convert.ToInt32(reiksmes[0]);
                                int uzd = Convert.ToInt32(reiksmes[1]);
                                WorkerInfo wi = new WorkerInfo();
                                wi.mxkd = Convert.ToDouble(Konvertuoti(reiksmes[2], ".", ","));
                                for (int j = 0; j < rod; j++)
                                {
                                    reiksmes = file.ReadLine().Split(',');
                                    for (int h = 0; h < reiksmes.Length; h++) reiksmes[h] = Konvertuoti(reiksmes[h], ".", ",");
                                    wi.Rodikliai.Add(new List<string>(reiksmes));
                                }
                                for (int j = 0; j < uzd; j++)
                                {
                                    reiksmes = file.ReadLine().Split(',');
                                    for (int h = 0; h < reiksmes.Length; h++) reiksmes[h] = Konvertuoti(reiksmes[h], ".", ",");
                                    wi.Uzduotys.Add(new List<string>(reiksmes));
                                }
                                workers.Add(wi);
                            }
                        }
                        else if (eilute == "Rodikliu lentele:")
                        {
                            int index = Convert.ToInt32(file.ReadLine());
                            for (int i = 0; i < index; i++)
                            {
                                string[] reiksmes = file.ReadLine().Split(',');
                                dataGridView2.Rows.Add(false, Konvertuoti(reiksmes[0], ".", ","), Convert.ToDouble(Konvertuoti(reiksmes[1], ".", ",")), Convert.ToDouble(Konvertuoti(reiksmes[2], ".", ",")), Convert.ToDouble(Konvertuoti(reiksmes[3], ".", ",")), Convert.ToDouble(Konvertuoti(reiksmes[4], ".", ","))); ;
                            }
                        }
                        else if (eilute == "Uzduociu lentele:")
                        {
                            int index = Convert.ToInt32(file.ReadLine());
                            for (int i = 0; i < index; i++)
                            {
                                string[] reiksmes = file.ReadLine().Split(',');
                                dataGridView3.Rows.Add(false, Konvertuoti(reiksmes[0], ".", ","), Convert.ToDouble(Konvertuoti(reiksmes[1], ".", ",")), Convert.ToDouble(Konvertuoti(reiksmes[2], ".", ",")));
                            }
                        }
                        else if (eilute == "Archyvo lentele:")
                        {
                            int index = Convert.ToInt32(file.ReadLine());
                            for (int i = 0; i < index; i++)
                            {
                                string eiluteSK = file.ReadLine();
                                string[] reiksmes = eiluteSK.Split(',');
                                for (int h = 0; h < reiksmes.Length; h++) reiksmes[h] = Konvertuoti(reiksmes[h], ".", ",");
                                Archyvas.Rows.Add(reiksmes[0], Convert.ToDouble(reiksmes[1]), Convert.ToDouble(reiksmes[2]), Convert.ToDouble(reiksmes[3]), Convert.ToDouble(reiksmes[4]), reiksmes[5], "( ? )");
                                reiksmes = file.ReadLine().Split(',');
                                int rod = Convert.ToInt32(reiksmes[0]);
                                int uzd = Convert.ToInt32(reiksmes[1]);
                                WorkerInfo wi = new WorkerInfo();
                                wi.mxkd = Convert.ToDouble(Konvertuoti(reiksmes[2],".",","));
                                for (int j = 0; j < rod; j++)
                                {
                                    reiksmes = file.ReadLine().Split(',');
                                    for (int h = 0; h < reiksmes.Length; h++) reiksmes[h] = Konvertuoti(reiksmes[h], ".", ",");
                                    wi.Rodikliai.Add(new List<string>(reiksmes));
                                }
                                for (int j = 0; j < uzd; j++)
                                {
                                    reiksmes = file.ReadLine().Split(',');
                                    for (int h = 0; h < reiksmes.Length; h++) reiksmes[h] = Konvertuoti(reiksmes[h], ".", ",");
                                    wi.Uzduotys.Add(new List<string>(reiksmes));
                                }
                                Arch_workers.Add(wi);
                            }
                        }
                    }
                }
                file.Close();
            }
        }
        public void LanguageNustatymas()
        {
            if (Alfredas.Properties.Settings.Default.Language == "English")
            {
                englishToolStripMenuItem.PerformClick();
            }
            if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
            {
                lietuviškaiToolStripMenuItem.PerformClick();
            }
            if (Alfredas.Properties.Settings.Default.Language == "Pусский")
            {
                русскийToolStripMenuItem.PerformClick();
            }
        }
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (FinishedLoading == true)
            {
                if (e.ColumnIndex == 0)
                {
                    if (List1CheckedItemIndex != -2)
                    {
                        if (List1CheckedItemIndex != e.RowIndex)
                        {
                            if (List1CheckedItemIndex != -1)
                            {
                                List1HasChanged = true;
                                dataGridView1.Rows[List1CheckedItemIndex].Cells[0].Value = false;
                            }
                            List1CheckedItemIndex = e.RowIndex;
                        }
                        else
                        {
                            if (List1HasChanged == true)
                            {
                                List1HasChanged = false;
                            }
                            else
                            {
                                List1CheckedItemIndex = -1;
                            }
                        }
                    }
                }
                else
                {
                    double rezult;
                    if (e.ColumnIndex == 2 && (dataGridView1.Rows[e.RowIndex].Cells[2].Value != null))
                    {
                        if(double.TryParse(dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString(), out rezult) == false)
                        {
                            if (Alfredas.Properties.Settings.Default.Language == "English")
                            {
                                MessageBox.Show("Write a correct number!");
                            }
                            if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                            {
                                MessageBox.Show("Įrašykite taisyklinga skaičių!");
                            }
                            if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                            {
                                MessageBox.Show("Написать правильный номер!"); // FIXME pataisyti vertima RUTAI
                            }
                            dataGridView1.Rows[e.RowIndex].Cells[2].Value = null;
                        }
                        else if (Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[2].Value) < 0)
                        {
                            if (Alfredas.Properties.Settings.Default.Language == "English")
                            {
                                MessageBox.Show("Write a positive number!");
                            }
                            if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                            {
                                MessageBox.Show("Įrašykite teigiama skaičių!");
                            }
                            if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                            {
                                MessageBox.Show("Введите положительное число!"); // FIXME pataisyti vertima RUTAI
                            }
                            dataGridView1.Rows[e.RowIndex].Cells[2].Value = null;
                            dataGridView1.Rows[e.RowIndex].Cells[5].Value = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[3].Value) + Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[4].Value);
                        }
                        else
                        {
                            dataGridView1.Rows[e.RowIndex].Cells[5].Value = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[2].Value) + Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[3].Value) + Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[4].Value);
                        }
                    }
                    CanExit = false;
                }
            }
        }
        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex == 0 && e.RowIndex != -1)
            {
                dataGridView1.EndEdit();
            }
        }
        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (dataGridView1.SelectedCells.Count == 1 && dataGridView1.SelectedCells[0].ColumnIndex == 0 && dataGridView1.SelectedCells[0].RowIndex != -1 && e.KeyValue == 32)
            {
                dataGridView1.EndEdit(); // removes functionality from spacebar
            }
        }
        private Tuple<double,double> calculateSalary() 
        {        
            
            // data1: 0 - boolean; 1 - vardas;      2 - Bazinis atlyginimas;    3 - Rodikliai;          4 - Uzduotys;           5 - Viso;
            // data2: 0 - boolean; 1 - pavadinimas; 2 - Bazine reiksme;         3 - Tiksline reiksme;    4 - Faktine reiksme;   5 - Maksimali kintama reiksme;
            // data3: 0 - boolean; 1 - pavadinimas; 2 - Maksimalus ivertinimas; 3 - Ivertinimas 
            double indicatorsSalary = 0;        
            double tasksSalary = 0;
            if(dataGridView2.Rows.Count != 0)
            {
                for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
                {
                    if (dataGridView2.Rows[i].Cells[0].Value != null)
                    {
                        if ((Boolean)dataGridView2.Rows[i].Cells[0].Value == true)
                        {
                            double TR = double.Parse(dataGridView2.Rows[i].Cells[3].Value.ToString());
                            double FR = double.Parse(dataGridView2.Rows[i].Cells[4].Value.ToString());
                            double BR = double.Parse(dataGridView2.Rows[i].Cells[2].Value.ToString());
                            double MKD = double.Parse(dataGridView2.Rows[i].Cells[5].Value.ToString());
                            if (TR > BR)
                            {
                                if ((BR <= FR) && (FR <= TR))
                                {
                                    indicatorsSalary += Math.Truncate(((FR - BR) / (TR - BR) * MKD) * 100) / 100;
                                }
                                if (FR > TR)
                                {
                                    indicatorsSalary += Math.Truncate(MKD * 100) / 100;
                                }
                            }
                            else if (TR < BR)
                            {
                                if ((TR <= FR) && (FR <= BR))
                                {
                                    indicatorsSalary += Math.Truncate(((FR - BR) / (TR - BR) * MKD) * 100) / 100;
                                }
                                if (FR < TR)
                                {
                                    indicatorsSalary += Math.Truncate(MKD * 100) / 100;
                                }
                            }
                        }
                    }
                }
            }
            double sfSum = 0;
            double sSum = 0;        
            if(dataGridView3.Rows.Count != 0)
            {
                for (int i = 0; i < dataGridView3.Rows.Count - 1; i++)
                {
                    if(dataGridView3.Rows[i].Cells[0].Value != null)
                    {
                        if ((Boolean)dataGridView3.Rows[i].Cells[0].Value == true)
                        {
                            sfSum += double.Parse(dataGridView3.Rows[i].Cells[3].Value.ToString()); // ivertinimas
                            sSum += double.Parse(dataGridView3.Rows[i].Cells[2].Value.ToString());   // maksimalus ivertinimas      
                        }   
                    }
                }        
            }
            if (sSum != 0)
            {
                tasksSalary = Math.Truncate(((sfSum / sSum) * Convert.ToDouble(textBox1.Text))*100)/100;      // bendras ivertinimas visoms,  reiksmem maxKDU rasti
            }
            else
            {
                tasksSalary = 0;
            }
            return new Tuple<double, double>(indicatorsSalary, tasksSalary);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (TikrintiArPazymeti() == true)
            {
                Tuple<double, double> result = calculateSalary();
                if (result.Item1 < 0 || result.Item2 < 0)
                {
                    if (Alfredas.Properties.Settings.Default.Language == "English")
                    {
                        MessageBox.Show("Calculated result cannot be a negative number, please check your indicators and task values!");                   
                    }
                    if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                    {
                        MessageBox.Show("Apskaičiuotas rezultatas negali būti neigiamas skaičius, patikrinkite rodiklių ir užduočių reikšmes!");                   
                    }
                    if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                    {
                        MessageBox.Show("Рассчитано результат не может быть отрицательным числом, пожалуйста, проверьте показатели и значение задача!"); // FIXME pataisyti vertima RUTAI
                    }
                }
                else
                {
                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {
                        if (dataGridView1.Rows[i].Cells[0].Value != null)
                        {
                            if ((Boolean)dataGridView1.Rows[i].Cells[0].Value == true)
                            {
                                if (dataGridView1.Rows[i].Cells[2].Value is DBNull)
                                {
                                    dataGridView1.Rows[i].Cells[2].Value = 0;
                                }
                                dataGridView1.Rows[i].Cells[3].Value = result.Item1;
                                dataGridView1.Rows[i].Cells[4].Value = result.Item2;
                                dataGridView1.Rows[i].Cells[5].Value = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) + result.Item1 + result.Item2;
                                WorkerInfo wi = new WorkerInfo();
                                if (textBox1.Text.Equals(""))  wi.mxkd = 0;
                                else wi.mxkd = Convert.ToDouble(textBox1.Text);
                                for (int j = 0; j < dataGridView2.Rows.Count - 1; j++)
                                {
                                    if (dataGridView2.Rows[j].Cells[0].Value != null)
                                    {
                                        if ((Boolean)dataGridView2.Rows[j].Cells[0].Value == true)
                                        {
                                            List<string> str = new List<string>(new string[] { dataGridView2.Rows[j].Cells[1].Value.ToString(), dataGridView2.Rows[j].Cells[2].Value.ToString(), dataGridView2.Rows[j].Cells[3].Value.ToString(), dataGridView2.Rows[j].Cells[4].Value.ToString(), dataGridView2.Rows[j].Cells[5].Value.ToString() });
                                            wi.Rodikliai.Add(str);
                                        }
                                    }
                                }
                                for (int j = 0; j < dataGridView3.Rows.Count - 1; j++)
                                {
                                    if (dataGridView3.Rows[j].Cells[0].Value != null)
                                    {
                                        if ((Boolean)dataGridView3.Rows[j].Cells[0].Value == true)
                                        {
                                            List<string> str = new List<string>(new string[] { dataGridView3.Rows[j].Cells[1].Value.ToString(), dataGridView3.Rows[j].Cells[2].Value.ToString(), dataGridView3.Rows[j].Cells[3].Value.ToString() });
                                            wi.Uzduotys.Add(str);
                                        }
                                    }
                                }
                                workers[i] = wi;
                            }
                        }
                    }
                }
            }
        }
        private Boolean TikrintiArPazymeti()
        {
            Boolean ArYraRodiklis = false;
            Boolean ArYraTask = false;
            Boolean rado = false;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value != null)
                {
                    if ((Boolean)dataGridView1.Rows[i].Cells[0].Value == true)
                    {
                        rado = true;
                    }
                }
            }
            if (rado == false)
            {
                return false;
            }

            rado = false;
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {
                if (dataGridView2.Rows[i].Cells[0].Value != null)
                {
                    if ((Boolean)dataGridView2.Rows[i].Cells[0].Value == true)
                    {
                        rado = true;
                        ArYraRodiklis = true;
                    }
                }
            }

            rado = false;
            for (int i = 0; i < dataGridView3.Rows.Count - 1; i++)
            {
                if (dataGridView3.Rows[i].Cells[0].Value != null)
                {
                    if ((Boolean)dataGridView3.Rows[i].Cells[0].Value == true)
                    {
                        rado = true;
                        ArYraTask = true;
                    }
                }
            }

            if (ArYraTask == true || ArYraRodiklis == true)
            {
                double rezult;
                if(ArYraTask == true)
                {
                    if (Double.TryParse(textBox1.Text, out rezult) == false)
                    {
                        if (Alfredas.Properties.Settings.Default.Language == "English")
                        {
                            MessageBox.Show("Write a correct number!");
                        }
                        if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                        {
                            MessageBox.Show("Įrašykite taisyklinga skaičių!");
                        }
                        if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                        {
                            MessageBox.Show("Написать правильный номер!"); // FIXME pataisyti vertima RUTAI
                        }
                        textBox1.Text = "0";
                        return false;
                    }
                    else
                    {
                        if (rezult < 0)
                        {
                            if (Alfredas.Properties.Settings.Default.Language == "English")
                            {
                                MessageBox.Show("Write a positive number!");
                            }
                            if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                            {
                                MessageBox.Show("Įrašykite teigiama skaičių!");
                            }
                            if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                            {
                                MessageBox.Show("Введите положительное число!"); // FIXME pataisyti vertima RUTAI
                            }
                            textBox1.Text = "0";
                            return false;
                        }
                    }
                }
                return true;
            }
            else
            {
                return false;
            }
        }
        private void lietuviškaiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FinishedLoading = false;

            archyvasToolStripMenuItem.Text = "Archyvas";
            apieToolStripMenuItem.Text = "Apie";
            pagalbaToolStripMenuItem.Text = "Pagalba";
            keistiKalbąToolStripMenuItem.Text = "Pasirinkti kalbą";
            saugotiExcelFormatuToolStripMenuItem.Text = "Eksportuoti į Excel";

            label1.Text = "Darbuotojai";
            label2.Text = "Rodikliai";
            richTextBox1.Text = "Maksimali kintama dalis pagal užduotis:";
            label4.Text = "Užduotys";
            button1.Text = "Skaičiuoti";
            button2.Text = "Saugoti";

            Pasirinkti_darbuotoja.HeaderText = "Pasirinkti darbuotoją";
            Vardas.HeaderText = "Vardas";
            Bazinis_atlyginimas.HeaderText = "Bazinis atlyginimas";
            Rodikliai.HeaderText = "Atlyginimas pagal rodiklius";
            Užduotys.HeaderText = "Atlyginimas pagal užduotis";
            Viso.HeaderText = "Viso";
            Info.HeaderText = "Info";

            Rodiklio_pasirinkimas.HeaderText = "Pasirinkti rodiklį";
            Pavadinimas.HeaderText = "Pavadinimas";
            Bazinė_reikšmė.HeaderText = "Bazinė reikšmė";
            Faktinė_reikšmė.HeaderText = "Faktinė reikšmė";
            Tikslinė_reikšmė.HeaderText = "Tikslinė reikšmė";
            Maksimali_kintama_dalis.HeaderText = "Maksimali kintama dalis";

            Užduoties_pasirinkimas.HeaderText = "Pasirinkti užduotį";
            PavadinimasUzd.HeaderText = "Pavadinimas";
            Maksimalus_įvertinimas.HeaderText = "Maksimalus įvertinimas";
            Įvertinimas.HeaderText = "Įvertinimas";

            FinishedLoading = true;
            Alfredas.Properties.Settings.Default.Language = "Lietuviu";
            Alfredas.Properties.Settings.Default.Save();
        }
        private void englishToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FinishedLoading = false;

            archyvasToolStripMenuItem.Text = "Archive";
            apieToolStripMenuItem.Text = "About";
            pagalbaToolStripMenuItem.Text = "Help";
            keistiKalbąToolStripMenuItem.Text = "Select language";
            saugotiExcelFormatuToolStripMenuItem.Text = "Export to Excel";

            label1.Text = "Employees";
            label2.Text = "Indicators";
            richTextBox1.Text = "Maximum task-related variable part:";
            label4.Text = "Tasks";
            button1.Text = "Calculate";
            button2.Text = "Save";

            Pasirinkti_darbuotoja.HeaderText = "Select an employee";
            Vardas.HeaderText = "Name";
            Bazinis_atlyginimas.HeaderText = "Basic Salary";
            Rodikliai.HeaderText = "Salary by the indicators";
            Užduotys.HeaderText = "Salary by the tasks";
            Viso.HeaderText = "Sum";
            Info.HeaderText = "Info";

            Rodiklio_pasirinkimas.HeaderText = "Select indicator";
            Pavadinimas.HeaderText = "Title";
            Bazinė_reikšmė.HeaderText = "Baseline value";
            Faktinė_reikšmė.HeaderText = "Actual outcome";
            Tikslinė_reikšmė.HeaderText = "Target value";
            Maksimali_kintama_dalis.HeaderText = "Maximum variable part";

            Užduoties_pasirinkimas.HeaderText = "Select task";
            PavadinimasUzd.HeaderText = "Title";
            Maksimalus_įvertinimas.HeaderText = "Maximum evaluation";
            Įvertinimas.HeaderText = "Evaluation";

            FinishedLoading = true;
            Alfredas.Properties.Settings.Default.Language = "English";
            Alfredas.Properties.Settings.Default.Save();
        }
        private void русскийToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FinishedLoading = false;

            archyvasToolStripMenuItem.Text = "Архив";
            apieToolStripMenuItem.Text = "Описание";
            pagalbaToolStripMenuItem.Text = "Помощь";
            keistiKalbąToolStripMenuItem.Text = "Выберите язык";
            saugotiExcelFormatuToolStripMenuItem.Text = "Экспортировать в Excel";

            label1.Text = "Cотрудники";
            label2.Text = "Показатели";
            richTextBox1.Text = "Максимальная переменная часть в соответствии с заданиями:";
            label4.Text = "Задании";
            button1.Text = "Рассчитать";
            button2.Text = "Сохранить";

            Pasirinkti_darbuotoja.HeaderText = "Выберите сотрудника";
            Vardas.HeaderText = "Имя";
            Bazinis_atlyginimas.HeaderText = "Баз. зарплата";
            Rodikliai.HeaderText = "Зарплата относительно показателей";
            Užduotys.HeaderText = "Зарплата относительно заданий";
            Viso.HeaderText = "Итого";
            Info.HeaderText = "Инфо";

            Rodiklio_pasirinkimas.HeaderText = "Выберите показатели";
            Pavadinimas.HeaderText = "Название";
            Bazinė_reikšmė.HeaderText = "Базовое значение";
            Faktinė_reikšmė.HeaderText = "Фактический результат";
            Tikslinė_reikšmė.HeaderText = "Целевое значение";
            Maksimali_kintama_dalis.HeaderText = "Максимальная переменная часть";

            Užduoties_pasirinkimas.HeaderText = "Выберите задачу";
            PavadinimasUzd.HeaderText = "Название";
            Maksimalus_įvertinimas.HeaderText = "Максимальная оценка";
            Įvertinimas.HeaderText = "Oценка";

            FinishedLoading = true;
            Alfredas.Properties.Settings.Default.Language = "Pусский";
            Alfredas.Properties.Settings.Default.Save();
        }
        private void pagalbaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (Form form = new Form())
            {
                form.Text = pagalbaToolStripMenuItem.Text;
                form.ClientSize = new System.Drawing.Size(1100, 500);
                form.MinimumSize = new System.Drawing.Size(1100, 500);
                form.VerticalScroll.Enabled = true;
                form.Icon = Alfredas.Properties.Resources.dd66325dba3d29ac;
                PictureBox Box = new PictureBox();
                Box.Dock = DockStyle.None;
                Box.Location = new System.Drawing.Point(4, 4);
                if (apieToolStripMenuItem.Text == "Apie")
                {
                    Box.Image = Alfredas.Properties.Resources.HelpLT;
                }
                else if (apieToolStripMenuItem.Text == "About")
                {
                    Box.Image = Alfredas.Properties.Resources.HelpEN;
                }
                else if (apieToolStripMenuItem.Text == "Описание")
                {
                    Box.Image = Alfredas.Properties.Resources.HelpRU;
                }
                Box.Size = new System.Drawing.Size(Box.Image.Size.Width, Box.Image.Size.Height);
                forBox = new Panel();
                Box.MouseEnter += new System.EventHandler(this.Box_MouseEnter);
                forBox.Dock = DockStyle.Fill;
                forBox.Location = new System.Drawing.Point(4, 4);
                forBox.AutoScroll = true;
                forBox.Controls.Add(Box);
                form.Controls.Add(forBox);
                form.ShowDialog();
            } 
        }
        private void apieToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (Form form = new Form())
            {
                form.Text = apieToolStripMenuItem.Text;
                form.ClientSize = new System.Drawing.Size(1000, 500);
                form.MinimumSize = new System.Drawing.Size(1000, 500);
                form.Icon = Alfredas.Properties.Resources.dd66325dba3d29ac;
                RichTextBox about = new RichTextBox();
                about.Dock = DockStyle.Fill;
                about.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F);
                about.Location = new System.Drawing.Point(4, 4);
                about.ReadOnly = true;
                about.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
                if (apieToolStripMenuItem.Text == "Apie")
                {
                    about.Text = Alfredas.Properties.Resources.AboutLT;
                }
                else if (apieToolStripMenuItem.Text == "About")
                {
                    about.Text = Alfredas.Properties.Resources.AboutEN;
                }
                else if (apieToolStripMenuItem.Text == "Описание")
                {
                    about.Text = Alfredas.Properties.Resources.AboutRU;
                }
                form.Controls.Add(about);
                form.ShowDialog();
            }
        }
        private void archyvasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (Form form = new Form())
            {
                form.Text = archyvasToolStripMenuItem.Text;
                form.ClientSize = new System.Drawing.Size(1000, 500);
                form.MinimumSize = new System.Drawing.Size(1000, 500);
                form.Icon = Alfredas.Properties.Resources.dd66325dba3d29ac;
                DataGridView ArchyvasShow = new DataGridView();
                ArchyvasShow.AllowUserToResizeRows = false;
                ArchyvasShow.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                ArchyvasShow.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                ArchyvasShow.Dock = DockStyle.Fill;
                ArchyvasShow.Location = new System.Drawing.Point(4, 4);
                ArchyvasShow.AllowUserToAddRows = false;
                ArchyvasShow.ReadOnly = true;
                ArchyvasShow.UserDeletingRow += new System.Windows.Forms.DataGridViewRowCancelEventHandler(this.ArchyvasShow_UserDeletingRow);
                ArchyvasShow.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.ArchyvasShow_CellContentClick);
                for (int i = 1; i < dataGridView1.Columns.Count - 1; i++)
                {
                    ArchyvasShow.Columns.Add(dataGridView1.Columns[i].Name, dataGridView1.Columns[i].HeaderText);
                }
                if (Alfredas.Properties.Settings.Default.Language == "English")
                {
                    ArchyvasShow.Columns.Add("Date", "Date");
                }
                if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                {
                    ArchyvasShow.Columns.Add("Data", "Data");
                }
                if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                {
                    ArchyvasShow.Columns.Add("Дата", "Дата");
                }
                for (int i = 0; i < ArchyvasShow.Columns.Count; i++)
                {
                    ArchyvasShow.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                }
                DataGridViewColumn col = new DataGridViewColumn();
                col.HeaderText = dataGridView1.Columns[dataGridView1.Columns.Count - 1].HeaderText;
                col.Name = dataGridView1.Columns[dataGridView1.Columns.Count - 1].Name;
                col.CellTemplate = dataGridView1.Columns[dataGridView1.Columns.Count - 1].CellTemplate;
                ArchyvasShow.Columns.Add(col);
                for (int i = 0; i < Archyvas.Rows.Count - 1; i++)
                {
                    ArchyvasShow.Rows.Add(Archyvas.Rows[i].Cells[0].Value, Archyvas.Rows[i].Cells[1].Value, Archyvas.Rows[i].Cells[2].Value, Archyvas.Rows[i].Cells[3].Value, Archyvas.Rows[i].Cells[4].Value, Archyvas.Rows[i].Cells[5].Value, Archyvas.Rows[i].Cells[6].Value);
                }
                form.Controls.Add(ArchyvasShow);
                form.ShowDialog();
                
            } 
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Boolean CanSave = true;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value != null)
                {
                    if ((Boolean)dataGridView1.Rows[i].Cells[0].Value == true)
                    {
                        Boolean galima = true;
                        for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                        {
                            if (dataGridView1.Rows[i].Cells[j].Value == null)
                            {
                                galima = false;
                            }
                        }
                        if (galima == true)
                        {
                            DateTime dateTime = DateTime.UtcNow.Date;
                            string data = dateTime.ToString("yyyy.MM.dd");
                            Archyvas.Rows.Add(dataGridView1.Rows[i].Cells[1].Value, Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value), Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value), Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value), Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value), data, dataGridView1.Rows[i].Cells[6].Value);
                            WorkerInfo wi = new WorkerInfo();
                            wi.mxkd = workers[i].mxkd;
                            wi.Rodikliai.AddRange(workers[i].Rodikliai);
                            wi.Uzduotys.AddRange(workers[i].Uzduotys);
                            Arch_workers.Add(wi);
                        }
                        else
                        {
                            if (Alfredas.Properties.Settings.Default.Language == "English")
                            {
                                MessageBox.Show("You have not filled all the archived employee's data!");
                            }
                            if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                            {
                                MessageBox.Show("Nebaigėte pildyti archyvuojamo darbuotojo duomenų!");
                            }
                            if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                            {
                                MessageBox.Show("Вы не заполнили все данные о заархивированном сотрудника!");
                            }
                            CanSave = false;
                        }
                    }
                }
            }
            if (CanSave == true)
            {
                SaveAll();
            }
        }
        private Boolean SaveAll()
        {
            string location = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\Duom.txt";
            location = location.Substring(6);
            System.IO.StreamWriter file = new System.IO.StreamWriter(location, false);
            Boolean galima = true;
            Boolean galimaIsjungtiPrograma = true;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 1; j < dataGridView1.Rows[i].Cells.Count; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value == null)
                    {
                        galima = false;
                    }
                }
            }
            if (galima == true)
            {
                file.WriteLine("Darbuotoju lentele:");
                file.WriteLine((dataGridView1.Rows.Count - 1).ToString());
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 1; j < dataGridView1.Rows[i].Cells.Count; j++)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Value == null)
                        {
                            galima = false;
                        }
                    }
                    file.WriteLine(Konvertuoti(dataGridView1.Rows[i].Cells[1].Value.ToString(), ",", ".") + "," + Konvertuoti(dataGridView1.Rows[i].Cells[2].Value.ToString(), ",", ".") + "," + Konvertuoti(dataGridView1.Rows[i].Cells[3].Value.ToString(), ",", ".") + "," + Konvertuoti(dataGridView1.Rows[i].Cells[4].Value.ToString(), ",", ".") + "," + Konvertuoti(dataGridView1.Rows[i].Cells[5].Value.ToString(),",","."));
                    file.WriteLine(workers[i].Rodikliai.Count.ToString() + "," + workers[i].Uzduotys.Count.ToString() + "," + Konvertuoti(workers[i].mxkd.ToString(), ",","."));
                    for (int j = 0; j < workers[i].Rodikliai.Count; j++)
                    {
                        file.WriteLine(Konvertuoti(workers[i].Rodikliai[j][0].ToString(),",",".") + "," + Konvertuoti(workers[i].Rodikliai[j][1].ToString(), ",", ".") + "," + Konvertuoti(workers[i].Rodikliai[j][2].ToString(), ",", ".") + "," + Konvertuoti(workers[i].Rodikliai[j][3].ToString(), ",", ".") + "," + Konvertuoti(workers[i].Rodikliai[j][4].ToString(), ",", "."));
                    }
                    for (int j = 0; j < workers[i].Uzduotys.Count; j++)
                    {
                        file.WriteLine(Konvertuoti(workers[i].Uzduotys[j][0].ToString(),",",".") + "," + Konvertuoti(workers[i].Uzduotys[j][1].ToString(), ",", ".") + "," + Konvertuoti(workers[i].Uzduotys[j][2].ToString(), ",", "."));
                    }
                }
            }
            else
            {
                if (Alfredas.Properties.Settings.Default.Language == "English")
                {
                    MessageBox.Show("Not all data in the table of employees have been filled!");
                }
                if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                {
                    MessageBox.Show("Darbuotojų lentelėje yra neužpildytų duomenų!");
                }
                if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                {
                    MessageBox.Show("Не все данные в таблице сотрудников были заполнены!");
                }
                file.WriteLine("Darbuotoju lentele:");
                file.WriteLine((0).ToString());
                galimaIsjungtiPrograma = false;
            }
            galima = true;
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {
                for (int j = 1; j < dataGridView2.Rows[i].Cells.Count; j++)
                {
                    if (dataGridView2.Rows[i].Cells[j].Value == null)
                    {
                        galima = false;
                    }
                }
            }
            if (galima == true)
            {
                file.WriteLine("Rodikliu lentele:");
                file.WriteLine((dataGridView2.Rows.Count - 1).ToString());
                for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
                {
                    file.WriteLine(Konvertuoti(dataGridView2.Rows[i].Cells[1].Value.ToString(),",",".") + "," + Konvertuoti(dataGridView2.Rows[i].Cells[2].Value.ToString(), ",", ".") + "," + Konvertuoti(dataGridView2.Rows[i].Cells[3].Value.ToString(), ",", ".") + "," + Konvertuoti(dataGridView2.Rows[i].Cells[4].Value.ToString(), ",", ".") + "," + Konvertuoti(dataGridView2.Rows[i].Cells[5].Value.ToString(), ",", "."));
                }
            }
            else
            {
                if (Alfredas.Properties.Settings.Default.Language == "English")
                {
                    MessageBox.Show("Not all data in the table of indicators have been filled!");
                }
                if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                {
                    MessageBox.Show("Rodiklių lentelėje yra neužpildytų duomenų!");
                }
                if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                {
                    MessageBox.Show("Не все данные в таблице показателей были заполнены!");
                }
                file.WriteLine("Rodikliu lentele:");
                file.WriteLine((0).ToString());
                galimaIsjungtiPrograma = false;
            }
            galima = true;
            for (int i = 0; i < dataGridView3.Rows.Count - 1; i++)
            {
                for (int j = 1; j < dataGridView3.Rows[i].Cells.Count; j++)
                {
                    if (dataGridView3.Rows[i].Cells[j].Value == null)
                    {
                        galima = false;
                    }
                }
            }
            if (galima == true)
            {
                file.WriteLine("Uzduociu lentele:");
                file.WriteLine((dataGridView3.Rows.Count - 1).ToString());
                for (int i = 0; i < dataGridView3.Rows.Count - 1; i++)
                {
                    file.WriteLine(Konvertuoti(dataGridView3.Rows[i].Cells[1].Value.ToString(),",",".") + "," + Konvertuoti(dataGridView3.Rows[i].Cells[2].Value.ToString(), ",", ".") + "," + Konvertuoti(dataGridView3.Rows[i].Cells[3].Value.ToString(), ",", "."));
                }
            }
            else
            {
                if (Alfredas.Properties.Settings.Default.Language == "English")
                {
                    MessageBox.Show("Not all data in the table of tasks have been filled!");
                }
                if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                {
                    MessageBox.Show("Užduočių lentelėje yra neužpildytų duomenų!");
                }
                if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                {
                    MessageBox.Show("Не все данные в таблице задач были заполнены!");
                }
                file.WriteLine("Uzduociu lentele:");
                file.WriteLine((0).ToString());
                galimaIsjungtiPrograma = false;
            }
            galima = true;
            for (int i = 0; i < Archyvas.Rows.Count - 1; i++)
            {
                for (int j = 0; j < Archyvas.Rows[i].Cells.Count; j++)
                {
                    if (Archyvas.Rows[i].Cells[j].Value == null)
                    {
                        galima = false;
                    }
                }
            }
            if (galima == true)
            {
                file.WriteLine("Archyvo lentele:");
                file.WriteLine((Archyvas.Rows.Count - 1).ToString());
                for (int i = 0; i < Archyvas.Rows.Count - 1; i++)
                {
                    file.WriteLine(Konvertuoti(Archyvas.Rows[i].Cells[0].Value.ToString(),",",".") + "," + Konvertuoti(Archyvas.Rows[i].Cells[1].Value.ToString(), ",", ".") + "," + Konvertuoti(Archyvas.Rows[i].Cells[2].Value.ToString(), ",", ".") + "," + Konvertuoti(Archyvas.Rows[i].Cells[3].Value.ToString(), ",", ".") + "," + Konvertuoti(Archyvas.Rows[i].Cells[4].Value.ToString(), ",", ".") + "," + Konvertuoti(Archyvas.Rows[i].Cells[5].Value.ToString(), ",", "."));
                    file.WriteLine(Arch_workers[i].Rodikliai.Count.ToString() + "," + Arch_workers[i].Uzduotys.Count.ToString() + "," + Konvertuoti(Arch_workers[i].mxkd.ToString(),",","."));
                    for (int j = 0; j < Arch_workers[i].Rodikliai.Count; j++)
                    {
                        file.WriteLine(Konvertuoti(Arch_workers[i].Rodikliai[j][0].ToString(), ",", ".") + "," + Konvertuoti(Arch_workers[i].Rodikliai[j][1].ToString(), ",", ".") + "," + Konvertuoti(Arch_workers[i].Rodikliai[j][2].ToString(), ",", ".") + "," + Konvertuoti(Arch_workers[i].Rodikliai[j][3].ToString(), ",", ".") + "," + Konvertuoti(Arch_workers[i].Rodikliai[j][4].ToString(),",","."));
                    }
                    for (int j = 0; j < Arch_workers[i].Uzduotys.Count; j++)
                    {
                        file.WriteLine(Konvertuoti(Arch_workers[i].Uzduotys[j][0].ToString(), ",", ".") + "," + Konvertuoti(Arch_workers[i].Uzduotys[j][1].ToString(), ",", ".") + "," + Konvertuoti(Arch_workers[i].Uzduotys[j][2].ToString(),",","."));
                    }
                }
            }
            else
            {
                if (Alfredas.Properties.Settings.Default.Language == "English")
                {
                    MessageBox.Show("Not all data in the archive table have been filled!");
                }
                if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                {
                    MessageBox.Show("Archyvo lentelėje yra neužpildytų duomenų!");
                }
                if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                {
                    MessageBox.Show("Не все данные в таблице архивных были заполнены!"); // FIXME pataisyti vertima RUTAI
                }
                file.WriteLine("Archyvo lentele:");
                file.WriteLine((0).ToString());
                galimaIsjungtiPrograma = false;
            }
            file.Close();
            CanExit = true;
            if (galimaIsjungtiPrograma == true)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private string Konvertuoti(string reiksme, string kaOut, string kaIn)
        {
            string naujaReiksme = reiksme;
            if (reiksme.Contains(kaOut) == true) naujaReiksme = reiksme.Replace(kaOut, kaIn);
            return naujaReiksme;
        }
        private Boolean TikrintiArReikiaKonvertuoti()
        {
            double test = 2.0 / 4.0;
            string testString = test.ToString();
            if (testString.Contains(",") == true) return true;
            else return false;
        }
        private void Box_MouseEnter(object sender, EventArgs e)
        {
            forBox.Focus();
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Boolean isjungti = true;
            if (CanExit == false)
            {
                e.Cancel = true;
                isjungti = SaveAll();
            }
            if (isjungti == true)
            {
                System.Windows.Forms.Application.Exit();
            }
            else
            {
                CanExit = false;
            }
        }
        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            double rezult;
            if (FinishedLoading == true)
            {
                if (e.ColumnIndex > 1)
                {
                    if (dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null && double.TryParse(dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), out rezult) == false)
                    {
                        dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = null;
                        if (Alfredas.Properties.Settings.Default.Language == "English")
                        {
                            MessageBox.Show("Write a correct number");
                        }
                        if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                        {
                            MessageBox.Show("Įrašykite taisyklinga skaičių");
                        }
                        if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                        {
                            MessageBox.Show("Написать правильный номер");
                        }
                    }
                    if (e.ColumnIndex == 2)
                    {
                        if (Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[2].Value) == Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[3].Value))
                        {
                            dataGridView2.Rows[e.RowIndex].Cells[2].Value = null;
                            if (Alfredas.Properties.Settings.Default.Language == "English")
                            {
                                MessageBox.Show("Baseline value can't be equal to target value");
                            }
                            if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                            {
                                MessageBox.Show("Bazinė reikšmė negali būti lygi tikslinei reikšmei");
                            }
                            if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                            {
                                MessageBox.Show("Базовое значение не может быть равным целевому значению");
                            }
                        }
                    }
                    else if (e.ColumnIndex == 3)
                    {
                        if (Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[2].Value) == Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[3].Value))
                        {
                            dataGridView2.Rows[e.RowIndex].Cells[3].Value = null;
                            if (Alfredas.Properties.Settings.Default.Language == "English")
                            {
                                MessageBox.Show("Baseline value can't be equal to target value");
                            }
                            if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                            {
                                MessageBox.Show("Bazinė reikšmė negali būti lygi tikslinei reikšmei");
                            }
                            if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                            {
                                MessageBox.Show("Базовое значение не может быть равным целевому значению");
                            }
                        }
                    }
                }
                if (Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[5].Value) < 0)
                {
                    if (Alfredas.Properties.Settings.Default.Language == "English")
                    {
                        MessageBox.Show("Write a positive number!");
                    }
                    if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                    {
                        MessageBox.Show("Įrašykite teigiama skaičių!");
                    }
                    if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                    {
                        MessageBox.Show("Введите положительное число!");
                    }
                    dataGridView2.Rows[e.RowIndex].Cells[5].Value = null;
                }
            }
            CanExit = false;
        }
        private void dataGridView3_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (FinishedLoading == true)
            {
                if (e.ColumnIndex > 1)
                {
                    double rezult;
                    if (dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
                    {
                        if (double.TryParse(dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), out rezult) == false)
                        {
                            dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = null;
                            if (Alfredas.Properties.Settings.Default.Language == "English")
                            {
                                MessageBox.Show("Write a correct number");
                            }
                            if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                            {
                                MessageBox.Show("Įrašykite taisyklinga skaičių");
                            }
                            if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                            {
                                MessageBox.Show("Написать правильный номер");
                            }
                        }
                        else if (Convert.ToDouble(dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value) < 0)
                        {
                            if (Alfredas.Properties.Settings.Default.Language == "English")
                            {
                                MessageBox.Show("Write a positive number!");
                            }
                            if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                            {
                                MessageBox.Show("Įrašykite teigiama skaičių!");
                            }
                            if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                            {
                                MessageBox.Show("Введите положительное число!"); // FIXME pataisyti vertima RUTAI
                            }
                            dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = null;
                        }
                        if (e.ColumnIndex == 2)
                        {
                            if (dataGridView3.Rows[e.RowIndex].Cells[3].Value != null)
                            {
                                if (Convert.ToInt32(dataGridView3.Rows[e.RowIndex].Cells[3].Value) > Convert.ToInt32(dataGridView3.Rows[e.RowIndex].Cells[2].Value))
                                {
                                    if (Alfredas.Properties.Settings.Default.Language == "English")
                                    {
                                        MessageBox.Show("Evaluation higher than maximum evaluation");
                                    }
                                    if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                                    {
                                        MessageBox.Show("Įvertinimas didesnis nei maksimalus įvertinimas");
                                    }
                                    if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                                    {
                                        MessageBox.Show("Оценка выше максимальной оценки"); // FIXME pataisyti vertima RUTAI
                                    }
                                    dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = null;
                                }
                            }
                        }
                        else
                        {
                            if (dataGridView3.Rows[e.RowIndex].Cells[2].Value != null)
                            {
                                if (Convert.ToInt32(dataGridView3.Rows[e.RowIndex].Cells[3].Value) > Convert.ToInt32(dataGridView3.Rows[e.RowIndex].Cells[2].Value))
                                {
                                    if (Alfredas.Properties.Settings.Default.Language == "English")
                                    {
                                        MessageBox.Show("Evaluation higher than maximum evaluation");
                                    }
                                    if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                                    {
                                        MessageBox.Show("Įvertinimas didesnis nei maksimalus įvertinimas");
                                    }
                                    if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                                    {
                                        MessageBox.Show("Оценка выше максимальной оценки"); // FIXME pataisyti vertima RUTAI
                                    }
                                    dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = null;
                                }
                            }
                        }
                    } 
                }
            }
            CanExit = false;
        }
        private void ArchyvasShow_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
           Archyvas.Rows.RemoveAt(e.Row.Index);
           Arch_workers.RemoveAt(e.Row.Index);
           CanExit = false;
        }
        private void dataGridView1_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            CanExit = false;
            workers.RemoveAt(e.Row.Index);
        }
        private void dataGridView2_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            CanExit = false;
        }
        private void dataGridView3_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            CanExit = false;
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 6 && e.RowIndex < dataGridView1.RowCount - 1)
            {
                using (Form form = new Form())
                {
                    if (Alfredas.Properties.Settings.Default.Language == "English")
                    {
                        form.Text = "Employee information";
                    }
                    if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                    {
                        form.Text = "Darbuotojo informacija";
                    }
                    if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                    {
                        form.Text = "Сотрудник информация";
                    }
                    form.ClientSize = new System.Drawing.Size(1000, 500);
                    form.MinimumSize = new System.Drawing.Size(1000, 500);
                    form.Icon = Alfredas.Properties.Resources.dd66325dba3d29ac;

                    TableLayoutPanel table = new TableLayoutPanel();
                    table.ColumnCount = 1;
                    table.RowCount = 4;
                    table.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 60F));
                    table.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 100F));
                    table.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 35F));
                    table.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 35F));
                    table.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 60F));
                    table.Dock = System.Windows.Forms.DockStyle.Fill;
                    table.GrowStyle = System.Windows.Forms.TableLayoutPanelGrowStyle.FixedSize;
                    table.Location = new System.Drawing.Point(0, 0);
                    table.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;

                    DataGridView Darbuot = new DataGridView();
                    Darbuot.AllowUserToResizeRows = false;
                    Darbuot.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                    Darbuot.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                    Darbuot.Dock = DockStyle.Fill;
                    Darbuot.Location = new System.Drawing.Point(4, 4);
                    Darbuot.AllowUserToAddRows = false;
                    Darbuot.ReadOnly = true;
                    Darbuot.AllowUserToDeleteRows = false;
                    Darbuot.AllowUserToResizeColumns = false;
                    for (int i = 1; i < dataGridView1.Columns.Count - 1; i++)
                    {
                        Darbuot.Columns.Add(dataGridView1.Columns[i].Name, dataGridView1.Columns[i].HeaderText);
                    }
                    Darbuot.Rows.Add(dataGridView1.Rows[e.RowIndex].Cells[1].Value, dataGridView1.Rows[e.RowIndex].Cells[2].Value, dataGridView1.Rows[e.RowIndex].Cells[3].Value, dataGridView1.Rows[e.RowIndex].Cells[4].Value, dataGridView1.Rows[e.RowIndex].Cells[5].Value);
                    table.Controls.Add(Darbuot, 0, 0);

                    DataGridView Rodikl = new DataGridView();
                    Rodikl.AllowUserToResizeRows = false;
                    Rodikl.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                    Rodikl.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                    Rodikl.Dock = DockStyle.Fill;
                    Rodikl.Location = new System.Drawing.Point(4, 4);
                    Rodikl.AllowUserToAddRows = false;
                    Rodikl.ReadOnly = true;
                    Rodikl.AllowUserToDeleteRows = false;
                    Rodikl.AllowUserToResizeColumns = false;
                    for (int i = 1; i < dataGridView2.Columns.Count; i++)
                    {
                        Rodikl.Columns.Add(dataGridView2.Columns[i].Name, dataGridView2.Columns[i].HeaderText);
                    }
                    for (int i = 0; i < workers[e.RowIndex].Rodikliai.Count; i++)
                    {
                        Rodikl.Rows.Add(workers[e.RowIndex].Rodikliai[i][0], Convert.ToDouble(workers[e.RowIndex].Rodikliai[i][1]), Convert.ToDouble(workers[e.RowIndex].Rodikliai[i][2]), Convert.ToDouble(workers[e.RowIndex].Rodikliai[i][3]), Convert.ToDouble(workers[e.RowIndex].Rodikliai[i][4]));
                    }
                    table.Controls.Add(Rodikl, 0, 1);

                    DataGridView Uzduot = new DataGridView();
                    Uzduot.AllowUserToResizeRows = false;
                    Uzduot.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                    Uzduot.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                    Uzduot.Dock = DockStyle.Fill;
                    Uzduot.Location = new System.Drawing.Point(4, 4);
                    Uzduot.AllowUserToAddRows = false;
                    Uzduot.ReadOnly = true;
                    Uzduot.AllowUserToDeleteRows = false;
                    Uzduot.AllowUserToResizeColumns = false;
                    for (int i = 1; i < dataGridView3.Columns.Count; i++)
                    {
                        Uzduot.Columns.Add(dataGridView3.Columns[i].Name, dataGridView3.Columns[i].HeaderText);
                    }
                    for (int i = 0; i < workers[e.RowIndex].Uzduotys.Count; i++)
                    {
                        Uzduot.Rows.Add(workers[e.RowIndex].Uzduotys[i][0], Convert.ToDouble(workers[e.RowIndex].Uzduotys[i][1]), Convert.ToDouble(workers[e.RowIndex].Uzduotys[i][2]));
                    }
                    table.Controls.Add(Uzduot, 0, 2);

                    TableLayoutPanel stable = new TableLayoutPanel();
                    stable.RowCount = 1;
                    stable.ColumnCount = 2;
                    stable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
                    stable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 200F));
                    stable.Dock = System.Windows.Forms.DockStyle.Fill;
                    stable.GrowStyle = System.Windows.Forms.TableLayoutPanelGrowStyle.FixedSize;
                    stable.Location = new System.Drawing.Point(0, 0);
                    stable.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;
                    
                    RichTextBox text = new RichTextBox();
                    text.ReadOnly = true;
                    text.Dock = System.Windows.Forms.DockStyle.Fill;
                    text.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F);
                    text.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None;
                    text.Location = new System.Drawing.Point(4, 4);
                    text.Text = richTextBox1.Text;
                    text.ReadOnly = true;
                    stable.Controls.Add(text, 0, 0);

                    RichTextBox text2 = new RichTextBox();
                    text2.ReadOnly = true;
                    text2.Dock = System.Windows.Forms.DockStyle.Fill;
                    text2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F);
                    text2.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None;
                    text2.Location = new System.Drawing.Point(4, 4);
                    text2.Text = workers[e.RowIndex].mxkd.ToString();                                                            // sutvarkyti  veliau su duomenim!!!
                    text2.ReadOnly = true;
                    stable.Controls.Add(text2, 1, 0);
                    table.Controls.Add(stable, 0, 3);
                    
                    form.Controls.Add(table);
                    form.ShowDialog();
                }
            }
        }
        private void ArchyvasShow_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 6)
            {
                using (Form form = new Form())
                {
                    if (Alfredas.Properties.Settings.Default.Language == "English")
                    {
                        form.Text = "Employee information";
                    }
                    if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                    {
                        form.Text = "Darbuotojo informacija";
                    }
                    if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                    {
                        form.Text = "Сотрудник информация";
                    }
                    form.ClientSize = new System.Drawing.Size(1000, 500);
                    form.MinimumSize = new System.Drawing.Size(1000, 500);
                    form.Icon = Alfredas.Properties.Resources.dd66325dba3d29ac;
                    
                    TableLayoutPanel table = new TableLayoutPanel();
                    table.ColumnCount = 1;
                    table.RowCount = 4;
                    table.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
                    table.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
                    table.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 35F));
                    table.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 35F));
                    table.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 60F));
                    table.Dock = System.Windows.Forms.DockStyle.Fill;
                    table.GrowStyle = System.Windows.Forms.TableLayoutPanelGrowStyle.FixedSize;
                    table.Location = new System.Drawing.Point(0, 0);
                    table.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;

                    DataGridView Darbuot = new DataGridView();
                    Darbuot.AllowUserToResizeRows = false;
                    Darbuot.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                    Darbuot.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                    Darbuot.Dock = DockStyle.Fill;
                    Darbuot.Location = new System.Drawing.Point(4, 4);
                    Darbuot.AllowUserToAddRows = false;
                    Darbuot.ReadOnly = true;
                    Darbuot.AllowUserToDeleteRows = false;
                    Darbuot.AllowUserToResizeColumns = false;
                    for (int i = 1; i < dataGridView1.Columns.Count - 1; i++)
                    {
                        Darbuot.Columns.Add(dataGridView1.Columns[i].Name, dataGridView1.Columns[i].HeaderText);
                    }
                    if (Alfredas.Properties.Settings.Default.Language == "English")
                    {
                        Darbuot.Columns.Add("Date", "Date");
                    }
                    if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                    {
                        Darbuot.Columns.Add("Data", "Data");
                    }
                    if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                    {
                        Darbuot.Columns.Add("Дата", "Дата");
                    } //FIXME_RUTA ???
                    Darbuot.Rows.Add(Archyvas.Rows[e.RowIndex].Cells[0].Value, Archyvas.Rows[e.RowIndex].Cells[1].Value, Archyvas.Rows[e.RowIndex].Cells[2].Value, Archyvas.Rows[e.RowIndex].Cells[3].Value, Archyvas.Rows[e.RowIndex].Cells[4].Value, Archyvas.Rows[e.RowIndex].Cells[5].Value);
                    table.Controls.Add(Darbuot, 0, 0);

                    DataGridView Rodikl = new DataGridView();
                    Rodikl.AllowUserToResizeRows = false;
                    Rodikl.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                    Rodikl.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                    Rodikl.Dock = DockStyle.Fill;
                    Rodikl.Location = new System.Drawing.Point(4, 4);
                    Rodikl.AllowUserToAddRows = false;
                    Rodikl.ReadOnly = true;
                    Rodikl.AllowUserToDeleteRows = false;
                    Rodikl.AllowUserToResizeColumns = false;
                    for (int i = 1; i < dataGridView2.Columns.Count; i++)
                    {
                        Rodikl.Columns.Add(dataGridView2.Columns[i].Name, dataGridView2.Columns[i].HeaderText);
                    }
                    for (int i = 0; i < Arch_workers[e.RowIndex].Rodikliai.Count; i++)
                    {
                        Rodikl.Rows.Add(Arch_workers[e.RowIndex].Rodikliai[i][0], Convert.ToDouble(Arch_workers[e.RowIndex].Rodikliai[i][1]), Convert.ToDouble(Arch_workers[e.RowIndex].Rodikliai[i][2]), Convert.ToDouble(Arch_workers[e.RowIndex].Rodikliai[i][3]), Convert.ToDouble(Arch_workers[e.RowIndex].Rodikliai[i][4]));
                    }
                    table.Controls.Add(Rodikl, 0, 1);
                    DataGridView Uzduot = new DataGridView();
                    Uzduot.AllowUserToResizeRows = false;
                    Uzduot.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                    Uzduot.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                    Uzduot.Dock = DockStyle.Fill;
                    Uzduot.Location = new System.Drawing.Point(4, 4);
                    Uzduot.AllowUserToAddRows = false;
                    Uzduot.ReadOnly = true;
                    Uzduot.AllowUserToDeleteRows = false;
                    Uzduot.AllowUserToResizeColumns = false;
                    for (int i = 1; i < dataGridView3.Columns.Count; i++)
                    {
                        Uzduot.Columns.Add(dataGridView3.Columns[i].Name, dataGridView3.Columns[i].HeaderText);
                    }
                    for (int i = 0; i < Arch_workers[e.RowIndex].Uzduotys.Count; i++)
                    {
                        Uzduot.Rows.Add(Arch_workers[e.RowIndex].Uzduotys[i][0], Convert.ToDouble(Arch_workers[e.RowIndex].Uzduotys[i][1]), Convert.ToDouble(Arch_workers[e.RowIndex].Uzduotys[i][2]));
                    }
                    table.Controls.Add(Uzduot, 0, 2);

                    TableLayoutPanel stable = new TableLayoutPanel();
                    stable.RowCount = 1;
                    stable.ColumnCount = 2;
                    stable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
                    stable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 200F));
                    stable.Dock = System.Windows.Forms.DockStyle.Fill;
                    stable.GrowStyle = System.Windows.Forms.TableLayoutPanelGrowStyle.FixedSize;
                    stable.Location = new System.Drawing.Point(0, 0);
                    stable.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;

                    RichTextBox text = new RichTextBox();
                    text.ReadOnly = true;
                    text.Dock = System.Windows.Forms.DockStyle.Fill;
                    text.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F);
                    text.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None;
                    text.Location = new System.Drawing.Point(4, 4);
                    text.Text = richTextBox1.Text;
                    text.ReadOnly = true;
                    stable.Controls.Add(text, 0, 0);

                    RichTextBox text2 = new RichTextBox();
                    text2.ReadOnly = true;
                    text2.Dock = System.Windows.Forms.DockStyle.Fill;
                    text2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F);
                    text2.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None;
                    text2.Location = new System.Drawing.Point(4, 4);
                    text2.Text = Arch_workers[e.RowIndex].mxkd.ToString();
                    text2.ReadOnly = true;
                    stable.Controls.Add(text2, 1, 0);
                    table.Controls.Add(stable, 0, 3);

                    form.Controls.Add(table);
                    form.ShowDialog();
                }
            }
        }
        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            if (FinishedLoading == true)
            {
                FilldataGridview1(e.RowIndex - 1);
                workers.Add(new WorkerInfo());
            }
            CanExit = false;
        }
        private void FilldataGridview1(int rowIndex)
        {
            for (int i = 3; i < dataGridView1.Rows[rowIndex].Cells.Count - 1; i++)
            {
                if (dataGridView1.Rows[rowIndex].Cells[i].Value == null)
                {
                    dataGridView1.Rows[rowIndex].Cells[i].Value = 0;
                }
            }
            dataGridView1.Rows[rowIndex].Cells[6].Value = "( ? )";
        }
        private void dataGridView2_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            CanExit = false;
        }
        private void dataGridView3_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            CanExit = false;
        }
        private void Export_To_Excel()
        {
            string location = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\Goals_Results_Salary.xlsx";
            if(System.IO.File.Exists(location) == false)
            {
                Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook = Excel.Workbooks.Add(XlSheetType.xlWorksheet);
                Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet = (Worksheet)Excel.ActiveSheet;
                Create_Header(ExcelWorkSheet);
                CreateWorkers(ExcelWorkSheet);
                ExcelWorkBook.SaveAs(location);
                Excel.Visible = true;
            }
            else
            {
                Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook = Excel.Workbooks.Open(location);
                Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet = ExcelWorkBook.Worksheets.get_Item("Sheet1");
                CreateWorkers(ExcelWorkSheet);
                ExcelWorkBook.Save();
                Excel.Visible = true;
            }
        }
        private void Create_Header(Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet)
        {
            ExcelWorkSheet.Range[ExcelWorkSheet.Cells[1, 1], ExcelWorkSheet.Cells[1, 2]].Merge();
            if (Alfredas.Properties.Settings.Default.Language == "English")
            {
                ExcelWorkSheet.Cells[1, 1] = "Employee";
                ExcelWorkSheet.Cells[1, 3].Value = "Indicators";
                ExcelWorkSheet.Cells[1, 4].Value = "Max. Indicators-related variable part";
                ExcelWorkSheet.Cells[1, 5].Value = "Baseline value";
                ExcelWorkSheet.Cells[1, 6].Value = "Target value";
                ExcelWorkSheet.Cells[1, 7].Value = "Actual outcome";
                ExcelWorkSheet.Cells[1, 8].Value = "Tasks";
                ExcelWorkSheet.Cells[1, 9].Value = "Max. tasks-related variable part";
                ExcelWorkSheet.Cells[1, 10].Value = "Maximum evaluation";
                ExcelWorkSheet.Cells[1, 11].Value = "Evaluation";
            }
            if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
            {
                ExcelWorkSheet.Cells[1, 1] = "Darbuotojas";
                ExcelWorkSheet.Cells[1, 3].Value = "Rodikliai";
                ExcelWorkSheet.Cells[1, 4].Value = "Maksimali kintama dalis pagal rodiklius"; // FIX
                ExcelWorkSheet.Cells[1, 5].Value = "Bazinė reikšmė";
                ExcelWorkSheet.Cells[1, 6].Value = "Tikslinė reikšmė";
                ExcelWorkSheet.Cells[1, 7].Value = "Faktinė reikšmė";
                ExcelWorkSheet.Cells[1, 8].Value = "Užduotys";
                ExcelWorkSheet.Cells[1, 9].Value = "Maksimali kintama dalis pagal užduotis";
                ExcelWorkSheet.Cells[1, 10].Value = "Maksimalus įvertinimas";
                ExcelWorkSheet.Cells[1, 11].Value = "Įvertinimas";
            }
            if (Alfredas.Properties.Settings.Default.Language == "Pусский")
            {
                ExcelWorkSheet.Cells[1, 1] = "Сотрудники";
                ExcelWorkSheet.Cells[1, 3].Value = "Показатели";
                ExcelWorkSheet.Cells[1, 4].Value = "Max. Часть относительно показателей";
                ExcelWorkSheet.Cells[1, 5].Value = "Базовое значение";
                ExcelWorkSheet.Cells[1, 6].Value = "Целевое значение";
                ExcelWorkSheet.Cells[1, 7].Value = "Фактическое значение";
                ExcelWorkSheet.Cells[1, 8].Value = "Задании";
                ExcelWorkSheet.Cells[1, 9].Value = "Max. Часть в соответствии с заданиями";
                ExcelWorkSheet.Cells[1, 10].Value = "Mаксимальная оценка";
                ExcelWorkSheet.Cells[1, 11].Value = "Оценка";
            }
            ExcelWorkSheet.Range[ExcelWorkSheet.Cells[1, 1], ExcelWorkSheet.Cells[1, 11]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            ExcelWorkSheet.Range[ExcelWorkSheet.Cells[1, 1], ExcelWorkSheet.Cells[1, 11]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            ExcelWorkSheet.Range[ExcelWorkSheet.Cells[1, 1], ExcelWorkSheet.Cells[1, 11]].Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            ExcelWorkSheet.Range[ExcelWorkSheet.Cells[1, 1], ExcelWorkSheet.Cells[1, 11]].WrapText = true;
        }
        private void CreateWorkers(Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet)
        {
            int y = 1; while(true)
            {
                string value = (string)(ExcelWorkSheet.Cells[y, 1] as Microsoft.Office.Interop.Excel.Range).Value;
                if (value == null) break;
                else y++;
            }
            int y_Now = y;
            int y_Start = y;
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value != null)
                {
                    if ((Boolean)dataGridView1.Rows[i].Cells[0].Value == true)
                    {
                        //ExcelWorkSheet.Range[ExcelWorkSheet.Cells[y_Now, 1], ExcelWorkSheet.Cells[y_Now, 2]].Merge();
                        ExcelWorkSheet.Cells[y_Now, 1] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                        ExcelWorkSheet.Cells[y_Now, 1].Font.Bold = true;
                        DateTime dateTime = DateTime.UtcNow.Date;
                        string data = dateTime.ToString("yyyy.MM.dd");
                        ExcelWorkSheet.Cells[y_Now, 2] = data;
                        if (Alfredas.Properties.Settings.Default.Language == "English")
                        {
                            ExcelWorkSheet.Cells[y_Now + 1, 1] = "Basic salary";
                            ExcelWorkSheet.Cells[y_Now + 2, 1] = "Indicators related part";
                            ExcelWorkSheet.Cells[y_Now + 3, 1] = "Tasks related part";
                            ExcelWorkSheet.Cells[y_Now + 4, 1] = "Total";
                        }
                        if (Alfredas.Properties.Settings.Default.Language == "Lietuviu")
                        {
                            ExcelWorkSheet.Cells[y_Now + 1, 1] = "Bazinis atlyginimas";                    
                            ExcelWorkSheet.Cells[y_Now + 2, 1] = "Kintama dalis pagal rodiklius";
                            ExcelWorkSheet.Cells[y_Now + 3, 1] = "Kintama dalis pagal užduotis";
                            ExcelWorkSheet.Cells[y_Now + 4, 1] = "Viso";
                        }
                        if (Alfredas.Properties.Settings.Default.Language == "Pусский")
                        {
                            ExcelWorkSheet.Cells[y_Now + 1, 1] = "Баз. Зарплата";
                            ExcelWorkSheet.Cells[y_Now + 2, 1] = "Перем. Часть-показатели";
                            ExcelWorkSheet.Cells[y_Now + 3, 1] = "Перем. Часть-задании";
                            ExcelWorkSheet.Cells[y_Now + 4, 1] = "Итого";
                        }
                        ExcelWorkSheet.Cells[y_Now + 1, 2] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                        ExcelWorkSheet.Cells[y_Now + 2, 2] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                        ExcelWorkSheet.Cells[y_Now + 3, 2] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                        ExcelWorkSheet.Cells[y_Now + 4, 2] = dataGridView1.Rows[i].Cells[5].Value.ToString();

                        for (int j = 0; j < workers[i].Rodikliai.Count; j++)
                        {
                            ExcelWorkSheet.Cells[y_Now + j, 3] = workers[i].Rodikliai[j][0];
                            ExcelWorkSheet.Cells[y_Now + j, 4] = workers[i].Rodikliai[j][4];
                            ExcelWorkSheet.Cells[y_Now + j, 5] = workers[i].Rodikliai[j][1];
                            ExcelWorkSheet.Cells[y_Now + j, 6] = workers[i].Rodikliai[j][2];
                            ExcelWorkSheet.Cells[y_Now + j, 7] = workers[i].Rodikliai[j][3];
                        }
                        for (int j = 0; j < workers[i].Uzduotys.Count; j++)
                        {
                            ExcelWorkSheet.Cells[y_Now + j, 8] = workers[i].Uzduotys[j][0];
                            ExcelWorkSheet.Cells[y_Now + j, 10] = workers[i].Uzduotys[j][1];
                            ExcelWorkSheet.Cells[y_Now + j, 11] = workers[i].Uzduotys[j][2];
                        }
                        int step = Math.Max(5, workers[i].Rodikliai.Count);
                        step = Math.Max(step, workers[i].Uzduotys.Count);

                        ExcelWorkSheet.Range[ExcelWorkSheet.Cells[y_Now, 9], ExcelWorkSheet.Cells[y_Now + step - 1, 9]].Merge();
                        if (workers[i].Uzduotys.Count != 0) ExcelWorkSheet.Cells[y_Now, 9] = workers[i].mxkd.ToString();
                        ExcelWorkSheet.Range[ExcelWorkSheet.Cells[y_Now, 9], ExcelWorkSheet.Cells[y_Now + step - 1, 9]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                        ExcelWorkSheet.Range[ExcelWorkSheet.Cells[y_Now, 9], ExcelWorkSheet.Cells[y_Now + step - 1, 9]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        ExcelWorkSheet.Range[ExcelWorkSheet.Cells[y_Now, 1], ExcelWorkSheet.Cells[y_Now, 2]].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                        ExcelWorkSheet.Range[ExcelWorkSheet.Cells[y_Now, 3], ExcelWorkSheet.Cells[y_Now + step - 1, 11]].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                        ExcelWorkSheet.Range[ExcelWorkSheet.Cells[y_Now, 3], ExcelWorkSheet.Cells[y_Now + step - 1, 11]].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                        ExcelWorkSheet.Range[ExcelWorkSheet.Cells[y_Now, 1], ExcelWorkSheet.Cells[y_Now + step - 1, 11]].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                        ExcelWorkSheet.Range[ExcelWorkSheet.Cells[y_Now, 3], ExcelWorkSheet.Cells[y_Now + step - 1, 11]].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                        ExcelWorkSheet.Range[ExcelWorkSheet.Cells[y_Now, 3], ExcelWorkSheet.Cells[y_Now + step - 1, 11]].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                        ExcelWorkSheet.Range[ExcelWorkSheet.Cells[y_Now, 3], ExcelWorkSheet.Cells[y_Now + step - 1, 11]].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                        ExcelWorkSheet.Range[ExcelWorkSheet.Cells[y_Now, 3], ExcelWorkSheet.Cells[y_Now + step - 1, 11]].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                        ExcelWorkSheet.Range[ExcelWorkSheet.Cells[y_Now + step - 1, 1], ExcelWorkSheet.Cells[y_Now + step - 1, 2]].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                        y_Now += step;
                    }
                }
            }
            ExcelWorkSheet.Range[ExcelWorkSheet.Cells[1, 1], ExcelWorkSheet.Cells[y_Now - 1, 11]].Columns.AutoFit();
        }
        private void saugotiExcelFormatuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(SaveAll() == true)
            {
                Export_To_Excel();
                button2.PerformClick();
                SaveAll();
            }
        }
    }
    public class WorkerInfo
    {
        public List<List<string>> Rodikliai = new List<List<string>>();
        public List<List<string>> Uzduotys = new List<List<string>>();
        public double mxkd = 0;
    }
}
