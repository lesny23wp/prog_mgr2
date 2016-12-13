using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System;  
using System.Collections;  
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Threading;



namespace wersja2_programmag
{

    public partial class Form1 : Form
    {
        string lokalizacja;
        private Thread trd;
        public Form1()
        {
            InitializeComponent();

        }
        private void button1_Click(object sender, EventArgs e)
        { 
            lokalizacja = textBox1.Text;
            
            this.dataGridView1.Columns[2].Visible = false;
            this.dataGridView1.Columns[3].Visible = false;
            this.dataGridView1.Columns[4].Visible = false;
            this.dataGridView1.Columns[5].Visible = false;
            this.dataGridView1.Columns[6].Visible = false;
            this.dataGridView1.Columns[7].Visible = false;
            this.dataGridView1.Columns[8].Visible = false;
            this.dataGridView1.Columns[9].Visible = false;
            this.dataGridView1.Columns[10].Visible = false;
            this.dataGridView1.Columns[11].Visible = false;
            this.dataGridView1.Columns[12].Visible = false;
            this.dataGridView1.Columns[13].Visible = false;
            this.dataGridView1.Columns[14].Visible = false;
            this.dataGridView1.Columns[15].Visible = false;
            this.dataGridView1.Columns[18].Visible = false;
            this.dataGridView1.Columns[19].Visible = false;


            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();

            pictureBox1.Visible = false;
            pictureBox2.Visible = false;
            label1.Visible = false;
            listBox1.Visible = false; 

            GC.Collect();
            GC.WaitForPendingFinalizers();



            dataGridView1.Visible = true;

            if (checkBox13.Checked)
            {
                
                this.dataGridView1.Columns[2].Visible = true;
                this.dataGridView1.Columns[3].Visible = true;
                this.dataGridView1.Columns[4].Visible = true;
                this.dataGridView1.Columns[5].Visible = true;
                this.dataGridView1.Columns[6].Visible = true;
                this.dataGridView1.Columns[7].Visible = true;
                this.dataGridView1.Columns[8].Visible = true;
                this.dataGridView1.Columns[9].Visible = true;
                this.dataGridView1.Columns[10].Visible = true;
                this.dataGridView1.Columns[11].Visible = true;
                this.dataGridView1.Columns[12].Visible = true;
                this.dataGridView1.Columns[13].Visible = true;
                this.dataGridView1.Columns[14].Visible = true;
                this.dataGridView1.Columns[15].Visible = true;
                this.dataGridView1.Columns[18].Visible = true;
                this.dataGridView1.Columns[19].Visible = true;
                
            }
            

            if (checkBox1.Checked)
            {

                this.dataGridView1.Columns[2].Visible = true;
            }

            if (checkBox2.Checked)
            {

                this.dataGridView1.Columns[3].Visible = true;
            }


            if (checkBox3.Checked)
            {

                this.dataGridView1.Columns[4].Visible = true;
            }

            if (checkBox4.Checked)
            {

                this.dataGridView1.Columns[5].Visible = true;
            }


            if (checkBox5.Checked)
            {

                this.dataGridView1.Columns[6].Visible = true;
            }


            if (checkBox6.Checked)
            {

                this.dataGridView1.Columns[7].Visible = true;
            }


            if (checkBox7.Checked)
            {

                this.dataGridView1.Columns[8].Visible = true;
            }


            if (checkBox8.Checked)
            {

                this.dataGridView1.Columns[9].Visible = true;
            }

            if (checkBox9.Checked)
            {

                this.dataGridView1.Columns[10].Visible = true;
            }

            if (checkBox10.Checked)
            {

                this.dataGridView1.Columns[11].Visible = true;
            }

            if (checkBox11.Checked)
            {

                this.dataGridView1.Columns[12].Visible = true;
            }


            if (checkBox14.Checked)
            {

                this.dataGridView1.Columns[13].Visible = true;
            }


            if (checkBox12.Checked)
            {

                this.dataGridView1.Columns[14].Visible = true;
            }

            if (checkBox15.Checked)
            {

               this.dataGridView1.Columns[15].Visible = true;
            }
            if (checkBox16.Checked)
            {

                this.dataGridView1.Columns[18].Visible = true;
            }
            if (checkBox17.Checked)
            {

                this.dataGridView1.Columns[19].Visible = true;
            }
            
            
            string queryString3 = "select Obszar from dane1 where Lp=1";

            
            string query_ilosc_wierszy = "select max(Lp) from dane1";
            string query_numer_kom = "	select \"numer komórki w tescie\" from dane1 where Lp=";
            string query_srednica_1 = "select Średnica_1 from dane1 where Lp=";
            string query_srednica_2 = "select Średnica_2 from dane1 where Lp=";
            string query_elongation = "select Elongation from dane1 where Lp=";
            string query_solidity = "select Solidity from dane1 where Lp=";
            string query_eccentricity = "select Eccentricity from dane1 where Lp=";
            string query_obszar = "select Obszar from dane1 where Lp=";
            string query_srednia_jas_kom = "select [Średnia jasność komórki] from dane1 where Lp=";
            string query_srednia_jas_tła = "select [Średnia jasność tła] from dane1 where Lp=";
            string query_IOD = "select IOD from dane1 where Lp=";
            string query_AIOD = "select AIOD from dane1 where Lp=";
            string query_PLOIDY = "select PLOIDY from dane1 where Lp=";
            string query_CCP = "select CCP from dane1 where Lp=";
            string query_naz = "select Foto_cale from dane1 where Lp=";
            string query_nazwisko = "select [Imię i Nazwisko] from dane1 where Lp=";
            string query_pesel = "select Pesel from dane1 where Lp=";
           // string query_foto2 = ""; 

            string serwer;
            serwer = textBox2.Text;

            string connectionString =
            "Data Source=" +serwer+";" +
            "Initial Catalog=matlab_dane;" +
            "Integrated Security=SSPI;";

            SqlConnection conn = new SqlConnection(connectionString);
            int ilosc_wierszy;

            conn.Open();
            SqlCommand comm = new SqlCommand(query_ilosc_wierszy, conn);
            ilosc_wierszy = Convert.ToInt32(comm.ExecuteScalar());

            conn.Close();


            int licznik = 1;

            while (licznik <= ilosc_wierszy)
            
            {

                int Lp;
                int numer_kom_test;
                float srednica_1;
                float srednica_2;
                float elongation;
                float solidity;
                float eccentricity;
                int obszar;
                float srednia_jas_kom;
                float srednia_jas_tła;
                float IOD;
                float AIOD;
                float PLOIDY;
                float CCP;
                float nazwa_d;
                string Nazwisko;
                string Pesell;

                try
                {

                    conn.Open();

                    SqlCommand comm2 = new SqlCommand(queryString3, conn);
                    Lp = Convert.ToInt32(comm2.ExecuteScalar());

                    SqlCommand comm3 = new SqlCommand(query_numer_kom + licznik, conn);
                    numer_kom_test = Convert.ToInt32(comm3.ExecuteScalar());

                    SqlCommand comm4 = new SqlCommand(query_srednica_1 + licznik, conn);
                    srednica_1 = Convert.ToSingle(comm4.ExecuteScalar());

                    SqlCommand comm5 = new SqlCommand(query_srednica_2 + licznik, conn);
                    srednica_2 = Convert.ToSingle(comm5.ExecuteScalar());

                    SqlCommand comm6 = new SqlCommand(query_elongation + licznik, conn);
                    elongation = Convert.ToSingle(comm6.ExecuteScalar());

                    SqlCommand comm7 = new SqlCommand(query_solidity + licznik, conn);
                    solidity = Convert.ToSingle(comm7.ExecuteScalar());

                    SqlCommand comm8 = new SqlCommand(query_eccentricity + licznik, conn);
                    eccentricity = Convert.ToSingle(comm8.ExecuteScalar());

                    SqlCommand comm9 = new SqlCommand(query_obszar + licznik, conn);
                    obszar = Convert.ToInt32(comm9.ExecuteScalar());

                    SqlCommand comm10 = new SqlCommand(query_srednia_jas_kom + licznik, conn);
                    srednia_jas_kom = Convert.ToSingle(comm10.ExecuteScalar());

                    SqlCommand comm11 = new SqlCommand(query_srednia_jas_tła + licznik, conn);
                    srednia_jas_tła = Convert.ToSingle(comm11.ExecuteScalar());

                    SqlCommand comm12 = new SqlCommand(query_IOD + licznik, conn);
                    IOD = Convert.ToSingle(comm12.ExecuteScalar());

                    SqlCommand comm13 = new SqlCommand(query_AIOD + licznik, conn);
                    AIOD = Convert.ToSingle(comm13.ExecuteScalar());

                    SqlCommand comm14 = new SqlCommand(query_PLOIDY + licznik, conn);
                    PLOIDY = Convert.ToSingle(comm14.ExecuteScalar());

                    SqlCommand comm15 = new SqlCommand(query_CCP + licznik, conn);
                    CCP = Convert.ToSingle(comm15.ExecuteScalar());

                    SqlCommand comm16 = new SqlCommand(query_naz + licznik, conn);
                    nazwa_d = Convert.ToSingle(comm16.ExecuteScalar());

                    SqlCommand comm17 = new SqlCommand(query_nazwisko + licznik, conn);
                    Nazwisko = Convert.ToString (comm17.ExecuteScalar());

                    SqlCommand comm18 = new SqlCommand(query_pesel + licznik, conn);
                    Pesell = Convert.ToString(comm18.ExecuteScalar());

                    int a = dataGridView1.Rows.Add();
                    dataGridView1.Rows[a].Cells[0].Value = licznik; //Lp
                    dataGridView1.Rows[a].Cells[1].Value = numer_kom_test; //numer komorki w tescie
                    dataGridView1.Rows[a].Cells[2].Value = srednica_1;//srednica_1
                    dataGridView1.Rows[a].Cells[3].Value = srednica_2;////srednica_2
                    dataGridView1.Rows[a].Cells[4].Value = elongation;//Elongation
                    dataGridView1.Rows[a].Cells[5].Value = solidity;//Solidity
                    dataGridView1.Rows[a].Cells[6].Value = eccentricity;//Eccentricity
                    dataGridView1.Rows[a].Cells[7].Value = obszar;//obszar;//obszar
                    dataGridView1.Rows[a].Cells[8].Value = srednia_jas_kom;//Śred jasno komorski
                    dataGridView1.Rows[a].Cells[9].Value = srednia_jas_tła;//średn jasn tła
                    dataGridView1.Rows[a].Cells[10].Value = IOD;//iod
                    dataGridView1.Rows[a].Cells[11].Value = AIOD;////aiod
                    dataGridView1.Rows[a].Cells[12].Value = PLOIDY;//ploidy

                    dataGridView1.Rows[a].Cells[13].Value = CCP;//ploidy

                    string aa = Convert.ToString(licznik);
                    string aaa = Convert.ToString(numer_kom_test);

                    string nazwa_pliku = aa + "_" + aaa + ".png";

                    dataGridView1.Rows[a].Cells[14].Value = Image.FromFile(lokalizacja+@"\komorki\" + nazwa_pliku);

                    string sciezka12 = lokalizacja + @"\komorki\" ;

                    dataGridView1.Rows[a].Cells[16].Value = (sciezka12 + nazwa_pliku);

                    string aaaa = "xxx";

                    string x = Convert.ToString(nazwa_d);

                    string nazwa_pliku2 = x + ".jpg";

                   dataGridView1.Rows[a].Cells[15].Value = Image.FromFile(lokalizacja + @"\zdjecia\" + nazwa_pliku2);

                   string sciezka12_2 = lokalizacja +@"\zdjecia\";

                   dataGridView1.Rows[a].Cells[16].Value = (sciezka12_2 + nazwa_pliku2);

                   dataGridView1.Rows[a].Cells[18].Value = Nazwisko;

                   dataGridView1.Rows[a].Cells[19].Value = Pesell;

                }

                catch (SqlException)
                {
                    
                }
                finally
                {
                    conn.Close();
                }

                licznik = licznik + 1;

            }


        }

        private void checkBox13_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox13.Checked)
            {
                checkBox1.Enabled = false;
                checkBox2.Enabled = false;
                checkBox3.Enabled = false;
                checkBox4.Enabled = false;
                checkBox5.Enabled = false;
                checkBox6.Enabled = false;
                checkBox7.Enabled = false;
                checkBox8.Enabled = false;
                checkBox9.Enabled = false;
                checkBox10.Enabled = false;
                checkBox11.Enabled = false;
                checkBox12.Enabled = false;
                checkBox14.Enabled = false;
                checkBox15.Enabled = false;
                checkBox16.Enabled = false;
                checkBox17.Enabled = false;


                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox14.Checked = false;
                checkBox15.Checked = false;
                checkBox16.Checked = false;
                checkBox17.Checked = false;


            }
            else
            {
                checkBox1.Enabled = true;
                checkBox2.Enabled = true;
                checkBox3.Enabled = true;
                checkBox4.Enabled = true;
                checkBox5.Enabled = true;
                checkBox6.Enabled = true;
                checkBox7.Enabled = true;
                checkBox8.Enabled = true;
                checkBox9.Enabled = true;
                checkBox10.Enabled = true;
                checkBox11.Enabled = true;
                checkBox12.Enabled = true;
                checkBox14.Enabled = true;
                checkBox15.Enabled = true;
                checkBox16.Enabled = true;
                checkBox17.Enabled = true;

            }

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

                 foreach (DataGridViewRow row in dataGridView1.SelectedRows) {
            
                pictureBox2.Visible = true;

                listBox1.Items.Clear();
                pictureBox1.Image = dataGridView1.CurrentRow.Cells[14].Value as Image;
                pictureBox2.Image = dataGridView1.CurrentRow.Cells[15].Value as Image;
                listBox1.Visible = true;
                label1.Visible = true;
                label1.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);
                //listBox1.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);
                 listBox1.Items.Add(" Numer komórki w teście: "+Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value));
                 listBox1.Items.Add(" Średnica_1: " + Convert.ToString(dataGridView1.CurrentRow.Cells[2].Value));
                 listBox1.Items.Add(" Średnica_2: " + Convert.ToString(dataGridView1.CurrentRow.Cells[3].Value));
                 listBox1.Items.Add(" Elongation: " + Convert.ToString(dataGridView1.CurrentRow.Cells[4].Value));
                 listBox1.Items.Add(" Solidity: " + Convert.ToString(dataGridView1.CurrentRow.Cells[5].Value));
                 listBox1.Items.Add(" Eccentricity: " + Convert.ToString(dataGridView1.CurrentRow.Cells[6].Value));
                 listBox1.Items.Add(" Obszar: " + Convert.ToString(dataGridView1.CurrentRow.Cells[7].Value));
                 listBox1.Items.Add(" Średnia jasność komórki: " + Convert.ToString(dataGridView1.CurrentRow.Cells[8].Value));
                 listBox1.Items.Add(" Średnia jasność tła: " + Convert.ToString(dataGridView1.CurrentRow.Cells[9].Value));
                 listBox1.Items.Add(" IOD Float: " + Convert.ToString(dataGridView1.CurrentRow.Cells[10].Value));
                 listBox1.Items.Add(" AIOD: " + Convert.ToString(dataGridView1.CurrentRow.Cells[11].Value));
                 listBox1.Items.Add(" PLOIDY: " + Convert.ToString(dataGridView1.CurrentRow.Cells[12].Value));
                 listBox1.Items.Add(" CCP: " + Convert.ToString(dataGridView1.CurrentRow.Cells[13].Value));

                 listBox1.Items.Add(" Imie i Nazw: " + Convert.ToString(dataGridView1.CurrentRow.Cells[18].Value));
                 listBox1.Items.Add(" Pesel: " + Convert.ToString(dataGridView1.CurrentRow.Cells[19].Value));


            }

            }

        private void button2_Click(object sender, System.EventArgs e)
        {
            Thread trd = new Thread(new ThreadStart(this.ThreadTask));
            trd.IsBackground = true;
            trd.Start();
          
        }
        private void ThreadTask()
        {
            Microsoft.Office.Interop.Excel._Application Excel = new Microsoft.Office.Interop.Excel.Application();

            Workbook wb = Excel.Workbooks.Add(XlSheetType.xlWorksheet);

            Worksheet ws = (Worksheet)Excel.ActiveSheet;
            Excel.Visible = false;

            ws.Cells[1, 1] = "Lp";
            ws.Cells[1, 2] = "Numer komórki w teście";
            ws.Cells[1, 3] = "Średnica 1";
            ws.Cells[1, 4] = "Średnica 2";
            ws.Cells[1, 5] = "Elongation";
            ws.Cells[1, 6] = "Solidity";
            ws.Cells[1, 7] = "Eccentricity";
            ws.Cells[1, 8] = "Oszar";
            ws.Cells[1, 9] = "Średnia jasność komórki";
            ws.Cells[1, 10] = "Średnia jasność tła";
            ws.Cells[1, 11] = "IOD";
            ws.Cells[1, 12] = "AIOD";
            ws.Cells[1, 13] = "PLOIDY";
            ws.Cells[1, 14] = "CCP";
            ws.Cells[1, 15] = "Imię i nazwisko";
            ws.Cells[1, 16] = "Pesel";


            for (int j = 2; j <= dataGridView1.Rows.Count + 1; j++)
            {
                for (int i = 1; i <= 20; i++)
                {
                    if (i == 15 || i == 16 || i == 17 || i == 18)
                    {

                    }
                    else if (i == 19 || i == 20)
                    {
                        ws.Cells[j, i - 4] = dataGridView1.Rows[j - 2].Cells[i - 1].Value;
                    }
                    else
                    {
                        ws.Cells[j, i] = dataGridView1.Rows[j - 2].Cells[i - 1].Value;
                    }

                }

            }

            var xlsfilename = "dane";
            wb.Close(true, xlsfilename);
             
        }

       
        private void button3_Click(object sender, System.EventArgs e)
        {
            FolderBrowserDialog folderbrowser1 = new FolderBrowserDialog();

            folderbrowser1.ShowDialog();

            textBox1.Text = folderbrowser1.SelectedPath.ToString();

            lokalizacja = textBox1.Text;
        }

        private void Form1_Load(object sender, System.EventArgs e)
        {

        }
       
        }
    }

