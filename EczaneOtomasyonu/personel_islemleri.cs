using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//sql veritabanı için gerekli kütüphaneyi entegre ediyoruz
using System.Data.SqlClient;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace EczaneOtomasyonu
{
    public partial class personel_islemleri : Form
    {
        public personel_islemleri()
        {
            InitializeComponent();
        }
        sqlbaglantisi bag = new sqlbaglantisi();// sql bağlantı classımızdan nesne oluşturduk ve bağlantı için kullanacağız
        SqlCommand kmt = new SqlCommand(); //sql ekleme silme güncelleme listeleme işlemleri için sqlcommand nesnesi oluşturduk
        DataSet dtst = new DataSet();//datagridviewlere sql serverdaki tabloları aktarmak için kullanıyoruz.

        public void sigortaDoldur()
        {
           
            kmt.Connection = bag.baglan(); //sql komutuna bağlantı oluşturduk
            kmt.CommandText = "Select * from sigortaDurumTablosu";
            SqlDataReader oku;
            oku = kmt.ExecuteReader();
            //sigortaların hepsini çekiyoruz
            while (oku.Read())
            {//combobox içerisine aktarıyoruz verileri
                comboBox1.Items.Add(oku[1].ToString());
            }

           
            oku.Dispose();
           // comboBox1.Sorted = true;

        }
        public void personelListele() //personelleri datagridviewe listeleyecek fonksiyonumuz
        {
            dtst.Clear();

            SqlDataAdapter adtr = new SqlDataAdapter("select * From personel", bag.baglan());
            adtr.Fill(dtst, "personel");
            dataGridView1.DataMember = "personel";
            dataGridView1.DataSource = dtst;
            adtr.Dispose();
            //datagridviewi listeledimiz verilerle doldurduk
            // seçilen satırın tamamını seçmesini sağladık ve aşşağıdaki kodlarla kolonlardaki başlıkları düzenledik
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Yaka Kart No";
            dataGridView1.Columns[2].HeaderText = "Adı Soyadı";
            dataGridView1.Columns[3].HeaderText = "TC Kimlik No";
            dataGridView1.Columns[4].HeaderText = "Doğum Tarihi";
            dataGridView1.Columns[5].HeaderText = "Adresi";
            dataGridView1.Columns[6].HeaderText = "Telefonu";
            dataGridView1.Columns[7].HeaderText = "Email";
            dataGridView1.Columns[8].HeaderText = "İşe Giriş Tarihi";
            dataGridView1.Columns[9].HeaderText = "Sigorta Girişi";
            //form üzerindeki textboxları ve comboboxları temizledik
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            comboBox1.Text = "";
        }
        private void button5_Click(object sender, EventArgs e)
        {
            Menu menu = new Menu(); //yeni menü form oluşturduk ve açilmasini sağladık mevcut formuda kapattık
            menu.Show();
            this.Hide();
        }
    
        private void button1_Click(object sender, EventArgs e) //ekleme işlemini yapan butonumuz
        {

            kmt.Connection = bag.baglan(); ;//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "INSERT INTO personel(yaka_kart,adi_soyadi,tc_kimlik,d_tarihi,adresi,telefonu,email,ise_giris,sigortasi) VALUES ('" + textBox1.Text + "' ,'" + textBox2.Text + "' ,'" + textBox3.Text + "' ,'" + textBox4.Text + "' ,'" + textBox5.Text + "' ,'" + textBox6.Text + "' ,'" + textBox7.Text + "' ,'" + textBox9.Text + "' ,'" + comboBox1.Text + "' )";
            kmt.ExecuteNonQuery(); //ekleme sql komutunu çalıştırdık ve başarılı mesaj verip gridwiewi tekrar doldurma fonksiyonunu çağırdık.
            MessageBox.Show("Personel Ekleme işlemi tamamlandı ! ");
            kmt.Dispose();
            personelListele();
        }

        private void personel_islemleri_Load(object sender, EventArgs e)
        {//oluşturdumuz fonksiyonları form yüklenirken çağırdık
            personelListele();
            sigortaDoldur();
            textBox1.Focus();
        }

        private void button4_Click(object sender, EventArgs e) //bu butonumuzda seçili datagridviewdeki id ye göre silme işlemi yaptık
        {

            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "DELETE from personel WHERE id = '" + dataGridView1.CurrentRow.Cells[0].Value.ToString() + "'";

            kmt.ExecuteNonQuery(); //oluşturdumuz silme sql komutunu burada çalıştırdık ve aşşağıda başarılı mesajını verip formu tekrar doldurma fonksiyonunu çağırdık
            kmt.Dispose();
            MessageBox.Show("Personel Silme işlemi tamamlandı ! ");
            dtst.Clear();
            personelListele();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            //sql tablosundaki verileri form üzerindeki textbox ve comboboxlara aktardık.
            textBox1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            textBox4.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            textBox5.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            textBox6.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            textBox7.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            textBox9.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            comboBox1.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            //datapickerdeki seçilen tarihi kısa formatta textbox9 da aktardık.
            textBox9.Text = dateTimePicker1.Value.ToShortDateString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //personel listele fonksiyonunu çağırdık
            personelListele();
        }

        private void button3_Click(object sender, EventArgs e)
        {

            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "update personel set yaka_kart='" + textBox1.Text + "',adi_soyadi='" + textBox2.Text + "',tc_kimlik='" + textBox3.Text + "',d_tarihi='" + textBox4.Text + "',adresi='" + textBox5.Text + "',telefonu='" + textBox6.Text + "',email='" + textBox7.Text + "', ise_giris='" + textBox9.Text + "',sigortasi='"+comboBox1.Text+"' where  id = '" + dataGridView1.CurrentRow.Cells[0].Value.ToString() + "'";
            kmt.ExecuteNonQuery();//güncelleme sql komutunu çalıştırdık aşşağıda başarılı meesajı verdik ve personel listele metodumuzu tekrar çalıştırdık
            kmt.Dispose();

            MessageBox.Show("Personel Güncelleme işlemi tamamlandı ! ");
            dtst.Clear();
            personelListele();
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            dtst.Clear();

            SqlDataAdapter adtr = new SqlDataAdapter("select * From personel where tc_kimlik LIKE '%" + textBox8.Text + "%'", bag.baglan());
            //burada arama işlemi yaptık ancak  normal aramalardan farklı olarak listelerken LIKE Kullandık bu yazdıklarımız eğer tc içerisinde var ise sonuç verecektir
            //bire bir de karşılaştırma yapmayacaktır.
            adtr.Fill(dtst, "personel");
            dataGridView1.DataMember = "personel";
            dataGridView1.DataSource = dtst;
            adtr.Dispose(); //datagridviewi listeledimiz verilerle doldurduk
                            // seçilen satırın tamamını seçmesini sağladık ve aşşağıdaki kodlarla kolonlardaki başlıkları düzenledik
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Yaka Kart No";
            dataGridView1.Columns[2].HeaderText = "Adı Soyadı";
            dataGridView1.Columns[3].HeaderText = "TC Kimlik No";
            dataGridView1.Columns[4].HeaderText = "Doğum Tarihi";
            dataGridView1.Columns[5].HeaderText = "Adresi";
            dataGridView1.Columns[6].HeaderText = "Telefonu";
            dataGridView1.Columns[7].HeaderText = "Email";
            dataGridView1.Columns[8].HeaderText = "İşe Giriş Tarihi";
            dataGridView1.Columns[9].HeaderText = "Sigorta Girişi";
            //form üzerindeki textbox ve comboboxları temizledik
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            comboBox1.Text = "";
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {


                bag.baglan();
                DialogResult cevap;
                cevap = MessageBox.Show("Personel Tablosunu da Boşaltılsın İstiyor Musunuz ? ", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (cevap == DialogResult.Yes)
                {

                    SqlCommand tablobosalt = new SqlCommand(" delete from personel", bag.baglan());
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
                ffff.Graphics.DrawString("Personel Tablosu", myFont, sbrush, 375, 65);
                ffff.Graphics.DrawLine(myPen, 50, 95, 770, 95); //çizgi çizdik.

                myFont = new System.Drawing.Font("Calibri", 6, FontStyle.Bold); //Detay başlığını yazacağımız için fontu kalın yapıp puntoyu 10 yaptık.
                ffff.Graphics.DrawString("Yaka Kart No", myFont, sbrush, 50, 110); //Detay başlığı
                ffff.Graphics.DrawString("Adı Soyadı", myFont, sbrush, 125, 110); //Detay başlığı
                ffff.Graphics.DrawString("TC Kimlik No", myFont, sbrush, 200, 110); // Detay başlığı
                ffff.Graphics.DrawString("Doğum Tarihi", myFont, sbrush, 275, 110); //Detay başlığı
                ffff.Graphics.DrawString("Adresi", myFont, sbrush, 350, 110); //Detay başlığı
                
                ffff.Graphics.DrawString("Telefonu", myFont, sbrush, 500, 110); //Detay başlığı
                ffff.Graphics.DrawString("Email", myFont, sbrush, 555, 110); //Detay başlığı
                ffff.Graphics.DrawString("İşe Giriş Tarihi", myFont, sbrush, 630, 110); //Detay başlığı
                ffff.Graphics.DrawString("Sigorta Girişi", myFont, sbrush, 725, 110); //Detay başlığı
                ffff.Graphics.DrawLine(myPen, 25, 125, 770, 125); //Çizgi çizdik.

                int y = 150; //y koordinatının yerini belirledik.(Verilerin yazılmaya başlanacağı yer)

                myFont = new System.Drawing.Font("Calibri", 6); //fontu 10 yaptık.

                int i = 0;//satır sayısı için değişken tanımladık.
                while (i <= dataGridView1.Rows.Count)//döngüyü son satırda sonlandıracağız.
                {
                    ffff.Graphics.DrawString(dataGridView1[1, i].Value.ToString(), myFont, sbrush, 50, y);//1.sütun
                    ffff.Graphics.DrawString(dataGridView1[2, i].Value.ToString(), myFont, sbrush, 125, y);//2.sütun
                    ffff.Graphics.DrawString(dataGridView1[3, i].Value.ToString(), myFont, sbrush, 200, y);//3.sütun
                    ffff.Graphics.DrawString(dataGridView1[4, i].Value.ToString(), myFont, sbrush, 275, y);//4.sütun
                    ffff.Graphics.DrawString(dataGridView1[5, i].Value.ToString(), myFont, sbrush, 350, y);//5.sütun
                    ffff.Graphics.DrawString(dataGridView1[6, i].Value.ToString(), myFont, sbrush, 500, y);//4.sütun
                    ffff.Graphics.DrawString(dataGridView1[7, i].Value.ToString(), myFont, sbrush, 555, y);//5.sütun
                    ffff.Graphics.DrawString(dataGridView1[8, i].Value.ToString(), myFont, sbrush, 630, y);//5.sütun
                    ffff.Graphics.DrawString(dataGridView1[9, i].Value.ToString(), myFont, sbrush, 725, y);//5.sütun

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
    }
}
