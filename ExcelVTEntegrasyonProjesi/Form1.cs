
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Data.SqlClient;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelVTEntegrasyonProjesi
{
    public partial class Form1 : Form
    {
        SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;Initial Catalog=ProjelerVT;Integrated Security=True");

        public Form1()
        {
            InitializeComponent();
        }

        private void btnReadData_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbook workbook = excelApp.Workbooks.Add(System.Reflection.Missing.Value);
            Excel.Worksheet sayfa1 = workbook.Sheets[1];

            string[] basliklar = { "Person No", "Name", "SurName", "Adres", "City" };
            Excel.Range range;
            for (int i = 0; i < basliklar.Length; i++)
            {
                range = sayfa1.Cells[1, (1 + i)];
                range.Value2 = basliklar[i];
            }


            try
            {
                con.Open();

                string sqlcumlesi = "SELECT personNo,pname,psurname,padres,pcity FROM Person ";
                SqlCommand cmd = new SqlCommand(sqlcumlesi, con);
                SqlDataReader sdr = cmd.ExecuteReader();
                int satir = 2;
                while (sdr.Read())
                {
                    string pno = sdr[0].ToString();
                    string pad = sdr[1].ToString();
                    string psoyad = sdr[2].ToString();
                    string padres = sdr[3].ToString();
                    string pcitey = sdr[4].ToString();

                    richTextBox1.Text = richTextBox1.Text + pno + "  " + pad + "  " + psoyad + "  " + padres + "  " + pcitey + "  " + Environment.NewLine;

                    range = sayfa1.Cells[satir, 1];
                    range.Value2 = pno;
                    range = sayfa1.Cells[satir, 2];
                    range.Value2 = pad;
                    range = sayfa1.Cells[satir, 3];
                    range.Value2 = psoyad;
                    range = sayfa1.Cells[satir, 4];
                    range.Value2 = padres;
                    range = sayfa1.Cells[satir, 5];
                    range.Value2 = pcitey;
                    satir++;

                }


            }
            catch (Exception)
            {
                MessageBox.Show("you have proplem with SQL order", "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                }
            }

        }

        private void btnReadExcel_Click(object sender, EventArgs e)
        {
            Excel.Application exlApp;
            Excel.Workbook exlWorkBook;
            Excel.Worksheet exlWorkSheet;
            Excel.Range range;
            int rCnt = 0;
            int cCnt = 0;
            exlApp = new Excel.Application();
            exlWorkBook = exlApp.Workbooks.Open(@"....");
            exlWorkSheet = (Excel.Worksheet)exlWorkBook.Worksheets.get_Item(1);
            range = exlWorkSheet.UsedRange;
            richTextBox2.Clear();

            for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
            {
                ArrayList list = new ArrayList();

                for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                {
                    string okunanhucre = Convert.ToString((range.Cells[rCnt, cCnt] as Excel.Range).Value2);

                    richTextBox2.Text = richTextBox2.Text + okunanhucre + "  ";
                    list.Add(okunanhucre);


                }
                richTextBox2.Text = richTextBox2.Text + Environment.NewLine;
              try
                {
                   
                    con.Open();

                    SqlCommand cm = new SqlCommand("ÝNSERT ÝNTO Person (personNo,pname,psurname,padres,pcity)" +
                                                     "VALUES (@personNo,@pname,@psurname,@padres,@pcity,)", con);

                    cm.Parameters.AddWithValue("@personNo", list[0]);
                    cm.Parameters.AddWithValue("@pname", list[1]);
                    cm.Parameters.AddWithValue("@psurname", list[2]);
                    cm.Parameters.AddWithValue("@padres", list[3]);
                    cm.Parameters.AddWithValue("@pcity", list[4]);
                    cm.ExecuteNonQuery();

                }
                catch(Exception ex)
                {
                    MessageBox.Show("when your Data writing something goes worong" + ex.ToString());
                }
                finally
                {
                    if (con != null)
                    {
                        con.Close();
                    }
                }


            }

            exlApp.Quit();
            ReleaseObject(exlWorkSheet);
            ReleaseObject(exlWorkBook);
            ReleaseObject(exlApp);


        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }


        }
    }
}
