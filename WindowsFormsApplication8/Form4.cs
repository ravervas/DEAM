using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication8
{
    public partial class Form4 : Form
    {
        Form1 parentForm;           
        public int num_column;
        public double Max_selection;
        public double Min_selection;
        public Form f3, f11;       
        public double N;
        public double N1;
        public double N2;
        public int N_start;
        public int num_partcle;
        public int Row;      
        public double N_now;
        public Form4(Form1 parent)
        {
            Program.f4 = this;
            InitializeComponent();
            parentForm = parent;
            
        }
        


        private void Form4_Load(object sender, EventArgs e)
        {
            
            textBox2.Text = Math.Round(parentForm.dii[0], 2).ToString();
            textBox3.Text = Math.Round(parentForm.dii[Program.f1.kol_count_di-1], 2).ToString();
            if (Program.f1.Eng==true)
            {
                this.Text = Program.f1.Histogram;
                label1.Text = "Enter the number of columns to plot the histogram";
                label2.Text = "Minimum size:";
                label3.Text = "Maximum size:";
                label4.Text = "Columns number:";
                button2.Text = "Cancel";
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            parentForm.f4.Close();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            num_column = int.Parse(textBox1.Text);
            if (num_column == 0)
            {
                if (Program.f1.Rus==true)
                MessageBox.Show("Число столбцов не задано! Попробуйте снова", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng==true)
                    MessageBox.Show("Columns number are not determined! Try again", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Math.Round(parentForm.dii[Program.f1.kol_count_di - 1], 0) > parentForm.dii[Program.f1.kol_count_di - 1])
                Max_selection = Math.Round(parentForm.dii[Program.f1.kol_count_di - 1], 0);
            else
                Max_selection = Math.Round(parentForm.dii[Program.f1.kol_count_di - 1], 0) + 1;
            if (Math.Round(parentForm.dii[0], 0) > parentForm.dii[0])
            Min_selection = Math.Round(parentForm.dii[0], 0)-1;
            else
                Min_selection = Math.Round(parentForm.dii[0], 0);
            N = (Max_selection - Min_selection) / num_column;// шаг                      
            if (Math.Round(N, 0)>=N)
            N = Math.Round(N, 0);
            else
                N = Math.Round(N, 0)+1;            
            N1 = N;
            N = N + Min_selection;                    
            

            if (Program.f1.f3_open)
                Program.f3.Close();

            f3 = new Form3();           
            f3.Show();
            Program.f3.Part = true;

           

            N_start = 0;
            Row = 0;           
            N_now = 0;
           
            Program.f3.dataGridView1.Rows.Add(num_column);           
            while (N_now < parentForm.dii[Program.f1.kol_count_di-1])
             {               
                num_partcle = 0;
                for (int i = N_start; i < Program.f1.kol_count_di; i++)
                {
                    if (parentForm.dii[i] < N)
                    {
                        num_partcle = num_partcle + 1;
                        N_start = i+1;
                        N_now = parentForm.dii[i];
                    }
                }
                Program.f3.dataGridView1.Rows[Row].Cells[1].Value = num_partcle;
                if (Row == 0)
                {
                    Program.f3.dataGridView1.Rows[Row].Cells[0].Value = Min_selection+ N1 / 2;
                    N2 = Min_selection + N1 / 2;
                }
                else
                {
                    Program.f3.dataGridView1.Rows[Row].Cells[0].Value = N2 + N1;
                    N2= N2 + N1;
                }
                N = N + N1;
                Row = Row + 1;
            }
            Program.f3.Sum = 0;

            Program.f3.N = Row;
            for (int i = 0; i < Program.f3.N; i++)
            {
                Program.f3.Sum = Program.f3.Sum + Convert.ToDouble(Program.f3.dataGridView1[1, i].Value); // сначал номер стобца, затем номер строки
            }           

            for (int i = 0; i < Program.f3.N; i++)
            {
                Program.f3.V = Convert.ToDouble(Program.f3.dataGridView1[1, i].Value) / Program.f3.Sum * 100;
                Program.f3.dataGridView1.Rows[i].Cells[2].Value = Convert.ToString(Math.Round(Program.f3.V, 2));
                Program.f3.dataGridView1.Rows[i].Cells[3].Value = Math.Round(Math.Log(Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value)), 2);
            }
            Program.f3.dataGridView1.Refresh();

            Program.f3.button2_Click(sender, e);

            
            parentForm.f4.Close();


        }

        private void textBox2_Enter(object sender, EventArgs e)
        {
            button3.Select();
        }

        private void textBox3_Enter(object sender, EventArgs e)
        {
            button3.Select();
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                e.Handled = true;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                num_column = int.Parse(textBox1.Text);
            }
            catch
            {
                return;
            }
            if (num_column < 2)
            {
                chart1.Series[0].Points.Clear();
                return;
            }
            if (Math.Round(parentForm.dii[Program.f1.kol_count_di - 1], 0) > parentForm.dii[Program.f1.kol_count_di - 1])
                Max_selection = Math.Round(parentForm.dii[Program.f1.kol_count_di - 1], 0);
            else
                Max_selection = Math.Round(parentForm.dii[Program.f1.kol_count_di - 1], 0) + 1;
            if (Math.Round(parentForm.dii[0], 0) > parentForm.dii[0])
                Min_selection = Math.Round(parentForm.dii[0], 0) - 1;
            else
                Min_selection = Math.Round(parentForm.dii[0], 0);
            N = (Max_selection - Min_selection) / num_column;// шаг                      
            if (Math.Round(N, 0) >= N)
                N = Math.Round(N, 0);
            else
                N = Math.Round(N, 0) + 1;
            N1 = N;
            N = N + Min_selection;
            

            N_start = 0;
            Row = 0;
            N_now = 0;
           
            dataGridView1.Rows.Add(num_column);
            while (N_now < parentForm.dii[Program.f1.kol_count_di - 1])
            {
                num_partcle = 0;
                for (int i = N_start; i < Program.f1.kol_count_di; i++)
                {
                    if (parentForm.dii[i] < N)
                    {
                        num_partcle = num_partcle + 1;
                        N_start = i + 1;
                        N_now = parentForm.dii[i];
                    }
                }
                dataGridView1.Rows[Row].Cells[1].Value = num_partcle;
                if (Row == 0)
                {
                    dataGridView1.Rows[Row].Cells[0].Value = Min_selection + N1 / 2;
                    N2 = Min_selection + N1 / 2;
                }
                else
                {
                    dataGridView1.Rows[Row].Cells[0].Value = N2 + N1;
                    N2 = N2 + N1;
                }
                N = N + N1;
                Row = Row + 1;
            }
            double Sum = 0;

            N = Row;
            for (int i = 0; i < N; i++)
            {
                Sum = Sum + Convert.ToDouble(dataGridView1[1, i].Value); // сначал номер стобца, затем номер строки
            }

            for (int i = 0; i < N; i++)
            {
                double V = Convert.ToDouble(dataGridView1[1, i].Value) / Sum * 100;
                dataGridView1.Rows[i].Cells[2].Value = Convert.ToString(Math.Round(V, 2));
            }
            dataGridView1.Refresh();
            chart1.Series[0].Points.Clear();
            

            for (int i = 0; i < N; i++)
            {
                double x = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                double y = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                chart1.Series[0].Points.AddXY(x, y);
                
            }
            chart1.ChartAreas[0].AxisY.Minimum = 0;
            
        }

        private void Form4_FormClosing(object sender, FormClosingEventArgs e)
        {
            Program.f1.form4 = false;
        }
    }
}
