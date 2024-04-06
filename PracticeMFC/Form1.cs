using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace PracticeMFC
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public static string TxtCon = @"Data Source=(LocalDb)\MSSQLLocalDB;Initial Catalog=PracticeMFC;Integrated Security=True";

        List<string> ListOzid = new List<string>();
        List<string> ListIntens = new List<string>();
        List<string> ListIspoln = new List<string>();
        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'practiceMFCDataSet.Request2' table. You can move, or remove it, as needed.

            Timer.Start();

            RefreshData();
            LblTimeRefresh.Text= DateTime.Now.ToString("dd/MM/yyyy HH:mm");

        }

        void RefreshData()
        {
            ListIntens.Clear();
            DateTime Dt = DateTime.Now;
            DateTime Fdt = DateTime.Parse(Dt.ToString("yyyy-MM-dd HH:mm"));

            SqlConnection Con = new SqlConnection(TxtCon);
            SqlCommand Cmd = new SqlCommand("select * from Request", Con);
            Con.Open();
            using (SqlDataReader Res = Cmd.ExecuteReader())
            {
                while (Res.Read())
                {
                    int requestId = int.Parse(Res["IdRequest"].ToString());
                    DateTime dateRequest = DateTime.Parse(Res["DateRequest"].ToString());


                    if (Fdt < DateTime.Parse(Res["DateRequest"].ToString()))
                    {
                        ListIntens.Add($"{Res["IdRequest"]}-{Res["DateRequest"]}");
                        CountRowsDB = int.Parse(Res["IdRequest"].ToString());
                    }
                    else if (dateRequest.Day == DateTime.Today.Day)
                    {
                        using (SqlConnection ConTwo = new SqlConnection(TxtCon))
                        {
                            SqlCommand command = new SqlCommand($"update Request set State='Old' where IdRequest = {requestId}", ConTwo);
                            ConTwo.Open();
                            command.ExecuteNonQuery();
                            ConTwo.Close();

                            ListIspoln.Add(Res["IdRequest"].ToString());
                        }
                    }
                    else
                    {
                        using (SqlConnection ConTwo = new SqlConnection(TxtCon))
                        {
                            SqlCommand command = new SqlCommand($"delete from Request where IdRequest = {requestId}", ConTwo);
                            ConTwo.Open();
                            command.ExecuteNonQuery();
                            ConTwo.Close();
                        }
                    }
                }
            }
            Con.Close();

            this.request2TableAdapter.Fill(this.practiceMFCDataSet.Request2);
        }

        private void BtnShowNotes_Click(object sender, EventArgs e)
        {
            PanelForAll.Visible = false;
            if (DgvNotes.Visible == false)
                DgvNotes.Visible = true;
            else DgvNotes.Visible = false;
        }

        private void BtnCondition_Click(object sender, EventArgs e)
        {
            if (this.Size == new Size(1057, 790))
            {
                label12.Visible = false;
                LblKefZagr.Visible = false;
            }

            DgvNotes.Visible= false;
            if (PanelForAll.Visible==false)
                PanelForAll.Visible = true;
            else PanelForAll.Visible = false;

            RefreshAll();
        }

        private void BtnManipulate_Click(object sender, EventArgs e)
        {
            if (BtnManipulate.Text == "Управление системой" && PanelForAll.Visible == true)
                BtnManipulate.Text = "Сохранить";
            else
            {
                using (SqlConnection Con = new SqlConnection(TxtCon))
                {
                    SqlCommand command = new SqlCommand($"update Valuess set n='{n}', t='{t}' where Id = 1", Con);
                    Con.Open();
                    command.ExecuteNonQuery();
                    Con.Close();
                }
                BtnManipulate.Text = "Управление системой";
            }

            if(PanelForAll.Visible == true && PanelSwitcher.Visible==false)
                PanelSwitcher.Visible = true;
            else PanelSwitcher.Visible = false;

            if (PanelForAll.Visible == true && PanelSwitch2.Visible == false)
                PanelSwitch2.Visible = true;
            else PanelSwitch2.Visible = false;
            RefreshAll();
        }

        double n = 0, t = 0;
        double KefZagruzki, Intensivnost, IntensivnostObsluz, SrednTimeOzid, SrednZayavok;

        private void Form1_Resize(object sender, EventArgs e)
        {

            if (this.Size != new Size(1057, 790))
            {
                label12.Visible = true;
                LblKefZagr.Visible = true;
            }

        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            LblTimeRefresh.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm");
            RefreshData();
            RefreshAll();
        }

        private void BtnToDock_Click(object sender, EventArgs e)
        {
            if (PanelForAll.Visible == true)
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                string filePath = (Application.StartupPath + "\\Report.xlsx");

                FileInfo file = new FileInfo(filePath);
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    worksheet.Cells[1, 1].Value = "Описание показателей";
                    worksheet.Cells[1, 2].Value = "Производственные данные";
                    worksheet.Cells[2, 1].Value = "Ожидают в очереди:";
                    worksheet.Cells[2, 2].Value = LblOzid.Text;
                    worksheet.Cells[3, 1].Value = "Количество записей:";
                    worksheet.Cells[3, 2].Value = LblZapis.Text;
                    worksheet.Cells[4, 1].Value = "Заявок исполнено:";
                    worksheet.Cells[4, 2].Value = LblIspoln.Text;
                    worksheet.Cells[5, 1].Value = "Работает окон:";
                    worksheet.Cells[5, 2].Value = LblOkna.Text;
                    worksheet.Cells[6, 1].Value = "Время обслуживания:";
                    worksheet.Cells[6, 2].Value = LblTimeOb.Text;
                    worksheet.Cells[7, 1].Value = "Пропускная способность:";
                    worksheet.Cells[7, 2].Value = LblPropusk.Text;
                    worksheet.Cells[8, 1].Value = "Интенсивность обслуживания:";
                    worksheet.Cells[8, 2].Value = LblIntenObs.Text;
                    worksheet.Cells[9, 1].Value = "Коэффициент загрузки:";
                    worksheet.Cells[9, 2].Value = LblKefZagr.Text;
                    worksheet.Cells[11, 1].Value = "Данные актуальны на:";
                    worksheet.Cells[11,2].Value= DateTime.Now;
                    worksheet.Cells[12, 1].Value = "Внимание!";
                    if(Double.Parse(LblKefZagr.Text)>=1)
                        worksheet.Cells[12, 2].Value = "Необходима оптимизация процессов";
                    else
                    worksheet.Cells[12, 2].Value = "Все процессы работают нормально";
                    if ((SrednZayavok / Intensivnost) > 5)
                        worksheet.Cells[13, 2].Value = "Время на ожидание превышает обслуживаемое!";


                    // Центировать текст
                    ExcelRange columnRange = worksheet.Cells[1, 1, worksheet.Dimension.Rows, 1];
                    for (int i = 2; i <= columnRange.Rows; i++)
                        worksheet.Cells[i, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    ExcelRange tableRange = worksheet.Cells[1, 1, columnRange.Rows-1, 2];

                    // Установка стиля границ для таблицы
                    tableRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    tableRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    tableRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    tableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    // Пример установки цвета границ
                    tableRange.Style.Border.Top.Color.SetColor(Color.Black);
                    tableRange.Style.Border.Left.Color.SetColor(Color.Black);
                    tableRange.Style.Border.Right.Color.SetColor(Color.Black);
                    tableRange.Style.Border.Bottom.Color.SetColor(Color.Black);


                    worksheet.Column(1).AutoFit();
                    worksheet.Column(2).AutoFit();

                    package.Save();

                    Process.Start(Application.StartupPath + "\\Report.xlsx");
                }
            }
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            DateTime Dt=DateTime.Now;
            SqlConnection Con = new SqlConnection(TxtCon);
            SqlCommand Cmd = new SqlCommand("select * from Request where State='New'", Con);
            Con.Open();
            SqlDataReader Res = Cmd.ExecuteReader();
            while(Res.Read())
                Dt=DateTime.Parse(Res["DateRequest"].ToString());
            Con.Close();

            SqlConnection Con1 = new SqlConnection(TxtCon);
            DateTime Fdt = DateTime.Parse(Dt.ToString("yyyy/MM/dd HH:mm"));
            SqlCommand Cmd1 = new SqlCommand($"insert into Request (IdRequest, FIO, Passport, SNILS, INN, AddresRegistry, DateRequest, State)" +
                $" values ('{CountRowsDB + 1}','123','4584 246125 24.10.2019','50033650925','628991979404','123', '{Fdt.AddMinutes(t)}','New')", Con1);
            Con1.Open();
            Cmd1.ExecuteNonQuery();
            Con1.Close();
            ListIntens.Add("1");
            this.request2TableAdapter.Fill(this.practiceMFCDataSet.Request2);
            RefreshAll();
            CountRowsDB++;
        }

        int CountRowsDB = 0;
        private void Form1_Shown(object sender, EventArgs e)
        {
            SqlConnection Con = new SqlConnection(TxtCon);
            SqlCommand Cmd = new SqlCommand("select * from Valuess", Con);
            Con.Open();
            SqlDataReader Res = Cmd.ExecuteReader();
            while (Res.Read())
            {
                n = int.Parse(Res["n"].ToString());
                    t= int.Parse(Res["t"].ToString());
            }
            Con.Close();
        }

        private void BtnFormulas_Click(object sender, EventArgs e)
        {
            FormFormulas Frm = new FormFormulas();
            Frm.ShowDialog();
        }

        void RefreshAll()
        {
            if(ListIntens.Count>n)
            LblOzid.Text = $"{ListIntens.Count - n} чел.";
            else LblOzid.Text = $"0 чел.";
            // Все записи из бд, где state = "New" и Date сегодня
            Intensivnost = ListIntens.Count;
            LblZapis.Text = Intensivnost.ToString();
            // Выводить из базы кол-во заявок, где state = "Old" и Date сегодня
            LblIspoln.Text = ListIspoln.Count.ToString();

            LblOkna.Text = n.ToString();
            LblTimeOb.Text = t.ToString() + " мин";

            if ((60.0 / t * n).ToString().Length>5)
                LblPropusk.Text = (60.0 / t * n).ToString().Substring(0,5) + " чел/ч";
            else LblPropusk.Text = 60.0 / t * n + " чел/ч";
            IntensivnostObsluz = 1.0 / t;
            if (IntensivnostObsluz.ToString().Length > 5)
                LblIntenObs.Text = IntensivnostObsluz.ToString().Substring(0, 5) + " чел/м";
            else LblIntenObs.Text = IntensivnostObsluz.ToString() + " чел/м";

            KefZagruzki = Intensivnost / (n * IntensivnostObsluz) / 60;
            if (KefZagruzki.ToString().Length > 5)
                LblKefZagr.Text = KefZagruzki.ToString().Substring(0, 5);
            else LblKefZagr.Text = KefZagruzki.ToString();

            SrednZayavok = (Intensivnost * Intensivnost) / n * (IntensivnostObsluz * (n * KefZagruzki));
        }



        private void BtnWinPlus_Click(object sender, EventArgs e)
        {
            LblOkna.Text = (int.Parse(LblOkna.Text)+1).ToString();
            n = int.Parse(LblOkna.Text);
            RefreshAll();
        }

        private void BtnWinMinus_Click(object sender, EventArgs e)
        {
            if (int.Parse(LblOkna.Text) != 0)
                LblOkna.Text = (int.Parse(LblOkna.Text) - 1).ToString();
            n = int.Parse(LblOkna.Text);
            RefreshAll();
        }

        private void BtnTimePlus_Click(object sender, EventArgs e)
        {
            LblTimeOb.Text= (int.Parse(LblTimeOb.Text.Split(' ')[0])+1).ToString() + " мин";
            t= int.Parse(LblTimeOb.Text.Split(' ')[0]);
            RefreshAll();
        }

        private void BtnTimeMinus_Click(object sender, EventArgs e)
        {
            if((int.Parse(LblTimeOb.Text.Split(' ')[0]))!=1)
                LblTimeOb.Text = (int.Parse(LblTimeOb.Text.Split(' ')[0]) -1).ToString() + " мин";
            t = int.Parse(LblTimeOb.Text.Split(' ')[0]);
            RefreshAll();
        }
    }
}
