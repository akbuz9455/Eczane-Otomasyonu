using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//sql veritabanı kullanma kütüphanesi ekledik
using System.Data.SqlClient;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace EczaneOtomasyonu
{
    public partial class AsiTakipFormu : Form
    {
        public AsiTakipFormu()
        {
            InitializeComponent();
        }
        sqlbaglantisi bag = new sqlbaglantisi();// sql bağlantı classımızdan nesne oluşturduk ve bağlantı için kullanacağız
        SqlCommand kmt = new SqlCommand(); //sql ekleme silme güncelleme listeleme işlemleri için sqlcommand nesnesi oluşturduk
        DataSet dtst = new DataSet();//datagridviewlere sql serverdaki tabloları aktarmak için kullanıyoruz.
        private void button7_Click(object sender, EventArgs e)
        {
            Menu menu = new Menu(); //yeni menü form oluşturduk ve açilmasini sağladık mevcut formuda kapattık
            menu.Show();
            this.Hide();
        }

        public void hastaTC()
        {

            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "SELECT DISTINCT * from hasta"; //tekrarlanmicak şekilde hastaları listeledik
            SqlDataReader oku;
            oku = kmt.ExecuteReader();
            while (oku.Read())
            {//hastaların tcleri tekrarlanmicak şekilde combobox1e aktardık
                comboBox1.Items.Add(oku[1].ToString());
            }

            oku.Dispose();
            comboBox1.Sorted = true;

        }
        public void personelTC()
        {

            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "SELECT DISTINCT * from personel"; //tekrarlamicak şekilde personelleri listeledik
            SqlDataReader oku;
            oku = kmt.ExecuteReader();
            while (oku.Read())
            {//personellerin tclerini comboboxlara string türüne çevirerek aktardık
                comboBox2.Items.Add(oku[3].ToString());
            }

            oku.Dispose();
            comboBox2.Sorted = true;

        }


        public void asiDoldur()
        {
            dtst.Clear();
            //dataseti temizledik
            SqlDataAdapter adtr = new SqlDataAdapter("select * From asiTablosu", bag.baglan());
            //aşı tablosunu listeledik
            adtr.Fill(dtst, "asiTablosu");
            dataGridView1.DataMember = "asiTablosu";
            dataGridView1.DataSource = dtst;
            adtr.Dispose();
            //datagridviewi doldurduk
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            // seçilen satırın tamamını seçmesini sağladık ve aşşağıdaki kodlarla kolonlardaki başlıkları düzenledik
            dataGridView1.BackgroundColor = Color.White;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Aşı Vurulan TC";
            dataGridView1.Columns[2].HeaderText = "Aşı Vuran TC";
            dataGridView1.Columns[3].HeaderText = "Aşı Adı";
            dataGridView1.Columns[4].HeaderText = "Etki Süresi";
            dataGridView1.Columns[5].HeaderText = "Etkisi";
            dataGridView1.Columns[6].HeaderText = "Aşı vurulma Tarihi";
           //form üzerindeki textbox ve comboboxları temizledik
            textBox1.Text = "";
            textBox2.Text = "";
            richTextBox1.Text = "";
            textBox9.Text = "";
            comboBox1.Text = "";
            comboBox2.Text = "";
           
        }
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {//datetimepickerdeki seçili tarihi kısa formatta textbox9'a aktardık.
            textBox9.Text = dateTimePicker1.Value.ToShortDateString();
        }

        private void AsiTakipFormu_Load(object sender, EventArgs e)
        {//oluşturdumuz fonksiyonları form yüklenirken çağırdık
            asiDoldur();
            hastaTC();
            personelTC();
        }

        private void button1_Click(object sender, EventArgs e)//ekleme işlemini yapan butonumuzun kodları
        {
            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "INSERT INTO asiTablosu(vurulanTC,vuranTC,asiAdi,etkiSuresi,etkisi,asiVurulmaTarihi) VALUES ('" + comboBox1.Text + "' ,'" + comboBox2.Text + "' ,'" + textBox1.Text + "' ,'" + textBox2.Text + "' ,'" + richTextBox1.Text + "' ,'" + textBox9.Text + "')";
            kmt.ExecuteNonQuery();
            //ekleme sql komutunu çalıştırdık ve başarılı mesaj verip gridwiewi tekrar doldurma fonksiyonunu çağırdık.
            MessageBox.Show(" Aşı Vurma işleminiz tamamlanmıştır GEÇMİŞ OLSUN ! ");
            kmt.Dispose();
            asiDoldur();
        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //comboboxtaki veri değiştiği vakit comboboxtaki tcye göre hastadan adsoyad bilgisi alıp label3e yazdırdık.
            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "SELECT DISTINCT * from hasta where tc_kimlik ='"+comboBox1.Text+"'"; //burada sql komutumuzda DISTINCT kullanma sebebimiz aynı tcde
            //birden fazla hasta kaydı açılabileceği için tek  birtanesini çeksin.
            SqlDataReader oku;
            oku = kmt.ExecuteReader();
            if (oku.Read())
            {
               label3.Text=oku[2].ToString();
            }

            oku.Dispose();
          
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {  //comboboxtaki veri değiştiği vakit comboboxtaki tcye göre personelden adsoyad bilgisi alıp label7e yazdırdık.
            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "SELECT DISTINCT * from personel where tc_kimlik ='" + comboBox2.Text + "'";
            SqlDataReader oku;
            oku = kmt.ExecuteReader();
            if (oku.Read())
            {
                label7.Text = oku[2].ToString();
            }

            oku.Dispose();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            dtst.Clear();
            //dataseti temizledik
            SqlDataAdapter adtr = new SqlDataAdapter("select * From asiTablosu where vurulanTC='"+comboBox1.Text+"'", bag.baglan());
            //tcye göre listeleme komutumuzu çalıştırdık
            adtr.Fill(dtst, "asiTablosu");
            dataGridView1.DataMember = "asiTablosu";
            dataGridView1.DataSource = dtst;
            adtr.Dispose();
            //tablolarımızı doldurduk
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            // seçilen satırın tamamını seçmesini sağladık ve aşşağıdaki kodlarla kolonlardaki başlıkları düzenledik
            dataGridView1.BackgroundColor = Color.White;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Aşı Vurulan TC";
            dataGridView1.Columns[2].HeaderText = "Aşı Vuran TC";
            dataGridView1.Columns[3].HeaderText = "Aşı Adı";
            dataGridView1.Columns[4].HeaderText = "Etki Süresi";
            dataGridView1.Columns[5].HeaderText = "Etkisi";
            dataGridView1.Columns[6].HeaderText = "Aşı vurulma Tarihi";
        }

        private void button6_Click(object sender, EventArgs e)
        {
            dtst.Clear();
            //dataseti temizledik
            SqlDataAdapter adtr = new SqlDataAdapter("select * From asiTablosu where vuranTC='" + comboBox2.Text + "'", bag.baglan());
            //tcye göre listeleme komutumuzu çalıştırdık
            adtr.Fill(dtst, "asiTablosu");
            dataGridView1.DataMember = "asiTablosu";
            dataGridView1.DataSource = dtst;
            adtr.Dispose();
            //tablolarımızı doldurduk
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            // seçilen satırın tamamını seçmesini sağladık ve aşşağıdaki kodlarla kolonlardaki başlıkları düzenledik
            dataGridView1.BackgroundColor = Color.White;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Aşı Vurulan TC";
            dataGridView1.Columns[2].HeaderText = "Aşı Vuran TC";
            dataGridView1.Columns[3].HeaderText = "Aşı Adı";
            dataGridView1.Columns[4].HeaderText = "Etki Süresi";
            dataGridView1.Columns[5].HeaderText = "Etkisi";
            dataGridView1.Columns[6].HeaderText = "Aşı vurulma Tarihi";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            asiDoldur();
        }

        private void button4_Click(object sender, EventArgs e) //bu butonumuzda seçili datagridviewdeki id ye göre silme işlemi yaptık
        {

            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "DELETE from asiTablosu WHERE id = '" + dataGridView1.CurrentRow.Cells[0].Value.ToString() + "'";

            kmt.ExecuteNonQuery(); //oluşturdumuz silme sql komutunu burada çalıştırdık ve aşşağıda başarılı mesajını verip formu tekrar doldurma fonksiyonunu çağırdık
            kmt.Dispose();
            MessageBox.Show("Aşı Kaydı Silme işlemi tamamlandı ! ");
            dtst.Clear();
            asiDoldur();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //sql tablosundaki verileri form üzerindeki textbox ve comboboxlara aktardık.

            comboBox1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            comboBox2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox1.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            textBox2.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            richTextBox1.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            textBox9.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
         
        }

        private void button3_Click(object sender, EventArgs e)
        {
            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "update asiTablosu set vurulanTC='" + comboBox1.Text + "',vuranTC='" + comboBox2.Text + "',asiAdi='" + textBox1.Text + "',etkiSuresi='" + textBox2.Text + "',etkisi='" + richTextBox1.Text + "',asiVurulmaTarihi='" + textBox9.Text   + "' where  id = '" + dataGridView1.CurrentRow.Cells[0].Value.ToString() + "'";
            kmt.ExecuteNonQuery();
            kmt.Dispose();

            MessageBox.Show("Aşı Bilgileri Güncelleme işlemi tamamlandı ! ");
            dtst.Clear();
            asiDoldur();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            //arama işlemi yaptık  textbox üzerine her bir kelime yazıldığında tekrar arayacaktır.
            dtst.Clear();
            //dataseti temizledik
            SqlDataAdapter adtr = new SqlDataAdapter("select * From asiTablosu where vurulanTC LIKE '%" + textBox3.Text + "%' ", bag.baglan());
            //burada arama işlemi yaptık ancak  normal aramalardan farklı olarak listelerken LIKE Kullandık bu yazdıklarımız eğer tc içerisinde var ise sonuç verecektir
            //bire bir de karşılaştırma yapmayacaktır.
            adtr.Fill(dtst, "asiTablosu");
            dataGridView1.DataMember = "asiTablosu";
            dataGridView1.DataSource = dtst;
            adtr.Dispose();

            //datagridviewi listeledimiz verilerle doldurduk
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            // seçilen satırın tamamını seçmesini sağladık ve aşşağıdaki kodlarla kolonlardaki başlıkları düzenledik
            dataGridView1.BackgroundColor = Color.White;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Aşı Vurulan TC";
            dataGridView1.Columns[2].HeaderText = "Aşı Vuran TC";
            dataGridView1.Columns[3].HeaderText = "Aşı Adı";
            dataGridView1.Columns[4].HeaderText = "Etki Süresi";
            dataGridView1.Columns[5].HeaderText = "Etkisi";
            dataGridView1.Columns[6].HeaderText = "Aşı vurulma Tarihi";
            //form üzerindeki textbox ve comboboxları temizledik
            textBox1.Text = "";
            textBox2.Text = "";
            richTextBox1.Text = "";
            textBox9.Text = "";
            comboBox1.Text = "";
            comboBox2.Text = "";

        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {


                bag.baglan();
                DialogResult cevap;
                cevap = MessageBox.Show("Aşı Tablosunu da Boşaltılsın İstiyor Musunuz ? ", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (cevap == DialogResult.Yes)
                {
                    
                    SqlCommand tablobosalt = new SqlCommand(" delete from asiTablosu", bag.baglan());
                    tablobosalt.ExecuteNonQuery();
                  
                    ///
                }

                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = true;
                object Missing = Type.Missing;
                Workbook workbook = excel.Workbooks.Add(Missing);
                Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
                int StartCol = 1;
                int StartRow = 1;
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    Range myRange = (Range)sheet1.Cells[StartRow, StartCol + j];
                    myRange.Value2 = dataGridView1.Columns[j].HeaderText;
                }
                StartRow++;
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {

                        Range myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                        myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                        myRange.Select();


                    }
                }

            }
            catch (Exception hata)
            {

                MessageBox.Show("Hata Aldınız" + hata.Message);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            pageSetupDialog1.ShowDialog();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            printPreviewDialog1.Document = printDocument1;
            printPreviewDialog1.ShowDialog();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            DialogResult pdr = printDialog1.ShowDialog();
            if (pdr == DialogResult.OK)
            {
                printDocument1.Print();
            }
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs ffff)
        {

            try
            {
                //ÇİZİM BAŞLANGICI
                System.Drawing.Font myFont = new System.Drawing.Font("Calibri", 7); //font oluşturduk
                SolidBrush sbrush = new SolidBrush(Color.Black);//fırça oluşturduk
                Pen myPen = new Pen(Color.Black); //kalem oluşturduk

                ffff.Graphics.DrawString("Düzenlenme Tarihi: " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString(), myFont, sbrush, 50, 25);
                ffff.Graphics.DrawLine(myPen, 50, 45, 770, 45); // Çizgi çizdik... 1. Kalem, 2. X, 3. Y Koordinatı, 4. Uzunluk, 5. BitişX
                //Aşı Vurulan TC , Aşı Vuran TC, Aşı Adı , Etki Süresi , Etkisi, Aşı Vurulma Tarihi
                myFont = new System.Drawing.Font("Calibri", 8, FontStyle.Bold);//Fatura başlığı yazacağımız için fontu kalın yaptık ve puntoyu büyütüp 15 yaptık.
                ffff.Graphics.DrawString("Aşı Tablosu", myFont, sbrush, 375, 65);
                ffff.Graphics.DrawLine(myPen, 50, 95, 770, 95); //çizgi çizdik.

                myFont = new System.Drawing.Font("Calibri", 6, FontStyle.Bold); //Detay başlığını yazacağımız için fontu kalın yapıp puntoyu 10 yaptık.
                ffff.Graphics.DrawString("Aşı Vurulan TC", myFont, sbrush, 25, 110); //Detay başlığı
                ffff.Graphics.DrawString("Aşı Vuran TC", myFont, sbrush, 125, 110); //Detay başlığı
                ffff.Graphics.DrawString("Aşı Adı", myFont, sbrush, 250, 110); // Detay başlığı
                ffff.Graphics.DrawString("Etki Süresi", myFont, sbrush, 350, 110); //Detay başlığı
                ffff.Graphics.DrawString("Etkisi", myFont, sbrush, 500, 110); //Detay başlığı
                ffff.Graphics.DrawString("Aşı Vurulma Tarihi", myFont, sbrush, 700, 110); //Detay başlığı
                ffff.Graphics.DrawLine(myPen, 25, 125, 770, 125); //Çizgi çizdik.

                int y = 150; //y koordinatının yerini belirledik.(Verilerin yazılmaya başlanacağı yer)

                myFont = new System.Drawing.Font("Calibri", 6); //fontu 10 yaptık.

                int i = 0;//satır sayısı için değişken tanımladık.
                while (i <= dataGridView1.Rows.Count)//döngüyü son satırda sonlandıracağız.
                {
                    ffff.Graphics.DrawString(dataGridView1[1, i].Value.ToString(), myFont, sbrush, 25, y);//1.sütun
                    ffff.Graphics.DrawString(dataGridView1[2, i].Value.ToString(), myFont, sbrush, 125, y);//2.sütun
                    ffff.Graphics.DrawString(dataGridView1[3, i].Value.ToString(), myFont, sbrush, 250, y);//3.sütun
                    ffff.Graphics.DrawString(dataGridView1[4, i].Value.ToString(), myFont, sbrush, 350, y);//4.sütun
                    ffff.Graphics.DrawString(dataGridView1[5, i].Value.ToString(), myFont, sbrush, 500, y);//5.sütun
                    ffff.Graphics.DrawString(dataGridView1[6, i].Value.ToString(), myFont, sbrush, 700, y);//5.sütun
                    y += 20; //y koordinatını arttırdık.
                    i += 1; //satır sayısını arttırdık

                    //yeni sayfaya geçme kontrolü
                    if (y > 1000)
                    {
                        ffff.Graphics.DrawString("(Devamı -->)", myFont, sbrush, 700, y + 50);
                        y = 50;
                        break; //burada yazdırma sınırına ulaştığımız için while döngüsünden çıkıyoruz
                               //çıktığımızda while baştan başlıyor i değişkeni değer almaya devam ediyor
                               //yazdırma yeni sayfada başlamış oluyor
                    }
                }
                //çoklu sayfa kontrolü
                if (i < dataGridView1.RowCount - 1)
                {
                    ffff.HasMorePages = true;
                }
                else
                {
                    ffff.HasMorePages = false;
                    i = 0;
                }
                StringFormat myStringFormat = new StringFormat();
                myStringFormat.Alignment = StringAlignment.Far;
            }
            catch
            {
            }

        }

        private void printPreviewDialog1_Load(object sender, EventArgs e)
        {

        }
    }
}
