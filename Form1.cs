using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace TestPoc
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public string tamYol;

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Excel Dosyası |*.xlsx";
            file.ShowDialog();
            tamYol = file.FileName;
            
            string baglantiAdresi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tamYol + ";Extended Properties='Excel 12.0;IMEX=1;'";
            OleDbConnection baglanti = new OleDbConnection(baglantiAdresi);
            OleDbCommand komut = new OleDbCommand("Select * From [" + "welcome" + "$]", baglanti);
            baglanti.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(komut);
            DataTable data = new DataTable();
            da.Fill(data);
            dataGridView1.DataSource = data;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string baglantiAdresi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tamYol + ";Extended Properties='Excel 12.0;IMEX=1;'";
            OleDbConnection baglanti = new OleDbConnection(baglantiAdresi);

            OleDbCommand komut = new OleDbCommand("Select * From [" + "welcome" + "$] Where Tarih like '%" + textBox2.Text.ToString() + "' ", baglanti);

            OleDbDataAdapter da = new OleDbDataAdapter(komut);

            DataTable data = new DataTable();
            da.Fill(data);
            int sayac = 0;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                sayac++;
            }
            textBox3.Text = sayac.ToString();


            //-----------------------------------------------------------------------------------------------------------------------
            //-------------- 2.yol. Bu yol datagrid üzerinden çalışmaktadır.Sadece exceldeki yıllara göre arama yapmaktadır.---------
            //-----------------------------------------------------------------------------------------------------------------------

            //DateTime date = new DateTime(2021, 01, 1);
            //DateTime date2 = new DateTime(2020, 01, 1);
            //DateTime date3 = new DateTime(2019, 01, 1);
            //DateTime date4 = new DateTime(2018, 01, 1);
            //DateTime date5 = new DateTime(2017, 01, 1);

            //int sayac2021 = 0, sayac2020 = 0, sayac2019 = 0, sayac2018 = 0, sayac2017 = 0;
            //string s = "Tarih";

            //for (int i = 0; i < dataGridView1.Columns.Count; i++)
            //{
            //    if (dataGridView1.Columns[i].HeaderText == s)
            //    {
            //        for (int j = 0; j < dataGridView1.Rows.Count - 1; j++)
            //        {
            //            foreach (DataGridViewCell cell in dataGridView1.Rows[j].Cells)
            //            {
            //                if (cell.Value != null)
            //                {
            //                    string yil = cell.Value.ToString();
            //                    DateTime.TryParse(yil, out DateTime sonyil);

            //                    //DateTime d2 = Convert.ToDateTime(cell.Value);

            //                    int durum = DateTime.Compare(date,sonyil);

            //                    if (sonyil.Year == date.Year)
            //                    {
            //                        sayac2021++;                                    
            //                    }
            //                    else if (sonyil.Year == date2.Year)
            //                    {
            //                        sayac2020++;                                   
            //                    }
            //                    else if (sonyil.Year == date3.Year)
            //                    {
            //                        sayac2019++;                                   
            //                    }
            //                    else if (sonyil.Year == date4.Year)
            //                    {
            //                        sayac2018++;                                    
            //                    }
            //                    else if (sonyil.Year == date5.Year)
            //                    {
            //                        sayac2017++;                                   
            //                    }
            //                }
            //            }
            //        }
            //    }
            //}

            //string arananyil=textBox2.Text.ToString();

            //if (arananyil=="2021")
            //{
            //    textBox3.Text = sayac2021.ToString();
            //}
            //else if (arananyil == "2020")
            //{
            //    textBox3.Text = sayac2020.ToString();
            //}
            //else if (arananyil == "2019")
            //{
            //    textBox3.Text = sayac2019.ToString();
            //}
            //else if (arananyil == "2018")
            //{
            //    textBox3.Text = sayac2018.ToString();
            //}
            //else if (arananyil == "2017")
            //{
            //    textBox3.Text = sayac2017.ToString();
            //}
            //else
            //{
            //    MessageBox.Show("Var olmayan bir yıl girdiniz. Lütfen yılı baştan giriniz");
            //}

        }


        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Excel Dosyaları";
            save.DefaultExt = "txt";
            save.Filter = "Text(*.txt)|.txt";
            save.ShowDialog();
            string dosya = save.FileName;

            using (TextWriter tw = new StreamWriter(dosya)) //dosya+ ".txt";
            {
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    tw.Write($"{dataGridView1.Columns[i].HeaderText.ToString()}");

                    if (i != dataGridView1.Columns.Count - 1)
                    {
                        tw.Write(";");
                    }

                }
                tw.WriteLine();

                int satirsayisi = dataGridView1.RowCount;

                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        tw.Write($"{dataGridView1.Rows[i].Cells[j].Value.ToString()}");

                        if (j != dataGridView1.Columns.Count - 1)
                        {
                            tw.Write(";");
                        }
                    }
                    tw.WriteLine();
                }
            }
                text_kayit.Visible = true;
                label_kayit.Visible = true;
                text_kayit.Text = dataGridView1.Rows.Count.ToString();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            text_kayit.Visible = false;
            label_kayit.Visible = false;
        }

    }
}
