using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//sql veritabanı için kütüphanemiz
using System.Data.SqlClient;
using System.IO;
using Microsoft.Office.Interop.Excel;
namespace EczaneOtomasyonu
{
    public partial class hastaKayit : Form
    {
        public hastaKayit()
        {
            InitializeComponent();
        }
        sqlbaglantisi bag = new sqlbaglantisi(); // sql bağlantı classımızdan nesne oluşturduk ve bağlantı için kullanacağız
        SqlCommand kmt = new SqlCommand(); //sql ekleme silme güncelleme listeleme işlemleri için sqlcommand nesnesi oluşturduk
        DataSet dtst = new DataSet();//datagridviewlere sql serverdaki tabloları aktarmak için kullanıyoruz.
        private void button5_Click(object sender, EventArgs e)
        {
            Menu menu = new Menu(); //yeni menü form oluşturduk ve açilmasini sağladık mevcut formuda kapattık
            menu.Show();
            this.Hide();
        }

        public void hastaDoldur()
        {
            dtst.Clear();
            //dataseti temizledik
            SqlDataAdapter adtr = new SqlDataAdapter("select * From hasta", bag.baglan());
            //hasta listeleme sql komutunu yazdık
            adtr.Fill(dtst, "hasta");
            dataGridView1.DataMember = "hasta";
            dataGridView1.DataSource = dtst;
            adtr.Dispose();
            //datagridviewi çekilen verilerle doldurduk
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            // seçilen satırın tamamını seçmesini sağladık ve aşşağıdaki kodlarla kolonlardaki başlıkları düzenledik
            dataGridView1.BackgroundColor = Color.White;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "TC Kimlik";
            dataGridView1.Columns[2].HeaderText = "Adı Soyadı";
            dataGridView1.Columns[3].HeaderText = "Sosyal Güvencesi";
            dataGridView1.Columns[4].HeaderText = "Adresi";
            dataGridView1.Columns[5].HeaderText = "Telefonu";
            dataGridView1.Columns[6].HeaderText = "İlaç Kullanımı";
            dataGridView1.Columns[7].HeaderText = "Kullanım Şekli";
            dataGridView1.Columns[8].HeaderText = "İlaç Barkod";
            //form üzerindeki textbox ve comboboxları temizledik
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "";
            comboBox4.Text = "";
        }

        public void ilacBarkodDoldur() //ilaç barkodlarını çekme fonksiyonumuz
        {

            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "Select * from ilac"; //ilaç listeleme sql kodumuzu yazdık
            SqlDataReader oku;
            oku = kmt.ExecuteReader();
            while (oku.Read())
            {//ilaç barkodlarını comboboxlara doldurduk
                comboBox4.Items.Add(oku[1].ToString());
            }

            oku.Dispose();
            comboBox4.Sorted = true;

        }
        public void kullanimAjDok()
        {
         
            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "Select * from kullanimActok";//kullanimActok listeleme sql kodumuzu yazdık
            SqlDataReader oku;
            oku = kmt.ExecuteReader();
            while (oku.Read())
            {// ac dok verileri ile comboboxlara doldurduk
                comboBox2.Items.Add(oku[1].ToString());
            }
            
            oku.Dispose();
            comboBox2.Sorted = true;

        }

        public void guvence()
        {
         
            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "Select * from guvence";//guvence listeleme sql kodumuzu yazdık
            SqlDataReader oku;
            oku = kmt.ExecuteReader();
            while (oku.Read())
            {//sosyal güvencelerini comboboxlara doldurduk
                comboBox1.Items.Add(oku[1].ToString());
            }

        
            oku.Dispose();
            comboBox1.Sorted = true;

        }
        public void kullanimVakti()
        {
           
            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "Select * from kullanimVakti";//kullanimVakti listeleme sql kodumuzu yazdık
            SqlDataReader oku;
            oku = kmt.ExecuteReader();
            while (oku.Read())
            {//kullanim Vakti verileri comboboxlara doldurduk
                comboBox3.Items.Add(oku[1].ToString());
            }

         
            oku.Dispose();
           comboBox3.Sorted = true;

        }

        private void button3_Click(object sender, EventArgs e) //güncelleme butonu komutları burada bulunuyor
        {
            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "update hasta set tc_kimlik='" + textBox1.Text + "',adi_soyadi='" + textBox2.Text + "',sosyal_güvencesi='" + comboBox1.Text + "',adresi='" + textBox3.Text + "',telefonu='" + textBox4.Text + "',ilac_kullanimi='" + comboBox2.Text + "',kullanim_sekli='" + comboBox3.Text + "',ilac_barkod='"+comboBox4.Text+"' where  id = '" + dataGridView1.CurrentRow.Cells[0].Value.ToString() + "'";
            kmt.ExecuteNonQuery(); //güncelleme işlemini yapar komutumuz ardından başarılı mesajı verip hastadoldur fonksiyonumuzu çağırır
            kmt.Dispose();

            MessageBox.Show("Hasta Bilgileri Güncelleme işlemi tamamlandı ! ");
            dtst.Clear();
            hastaDoldur();
        }

        private void button1_Click(object sender, EventArgs e)//ekleme  işlemi için butonumuz.
        {
          
           kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "INSERT INTO hasta(tc_kimlik,adi_soyadi,sosyal_güvencesi,adresi,telefonu,ilac_kullanimi,kullanim_sekli,ilac_barkod) VALUES ('" + textBox1.Text + "' ,'" + textBox2.Text + "' ,'" + comboBox1.Text + "' ,'" + textBox3.Text + "' ,'" + textBox4.Text + "' ,'" + comboBox2.Text + "' ,'" + comboBox3.Text + "','"+comboBox4.Text+"')";
            kmt.ExecuteNonQuery();
            //ekleme sql komutunu çalıştırdık ve başarılı mesaj verip gridwiewi tekrar doldurma fonksiyonunu çağırdık.
            MessageBox.Show(" kayıt işleminiz tamamlanmıştır GEÇMİŞ OLSUN ! ");
           kmt.Dispose();
            hastaDoldur();
        }

        private void hastaKayit_Load(object sender, EventArgs e)
        {//oluşturdumuz fonksiyonları form yüklenirken çağırdık
            hastaDoldur();
            guvence();
            ilacBarkodDoldur();
            kullanimAjDok();
            kullanimVakti();
        }

        private void button4_Click(object sender, EventArgs e) //bu butonumuzda seçili datagridviewdeki id ye göre silme işlemi yaptık
        {

            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "DELETE from hasta WHERE id = '" + dataGridView1.CurrentRow.Cells[0].Value.ToString() + "'";

            kmt.ExecuteNonQuery(); //oluşturdumuz silme sql komutunu burada çalıştırdık ve aşşağıda başarılı mesajını verip formu tekrar doldurma fonksiyonunu çağırdık
            kmt.Dispose();
            MessageBox.Show("Hasta Kaydı Silme işlemi tamamlandı ! ");
            dtst.Clear();
            hastaDoldur();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        { //sql tablosundaki verileri form üzerindeki textbox ve comboboxlara aktardık.

            textBox1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            comboBox1.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
         
            textBox4.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            comboBox2.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            comboBox3.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            comboBox4.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            dtst.Clear();

            SqlDataAdapter adtr = new SqlDataAdapter("select * From hasta where tc_kimlik LIKE '%"+textBox5.Text+"%'", bag.baglan());
            //burada arama işlemi yaptık ancak  normal aramalardan farklı olarak listelerken LIKE Kullandık bu yazdıklarımız eğer tc içerisinde var ise sonuç verecektir
            //bire bir de karşılaştırma yapmayacaktır.
            adtr.Fill(dtst, "hasta");
            dataGridView1.DataMember = "hasta";
            dataGridView1.DataSource = dtst;
            adtr.Dispose();
            //datagridviewi listeledimiz verilerle doldurduk
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            // seçilen satırın tamamını seçmesini sağladık ve aşşağıdaki kodlarla kolonlardaki başlıkları düzenledik
            dataGridView1.BackgroundColor = Color.White;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "TC Kimlik";
            dataGridView1.Columns[2].HeaderText = "Adı Soyadı";
            dataGridView1.Columns[3].HeaderText = "Sosyal Güvencesi";
            dataGridView1.Columns[4].HeaderText = "Adresi";
            dataGridView1.Columns[5].HeaderText = "Telefonu";
            dataGridView1.Columns[6].HeaderText = "İlaç Kullanımı";
            dataGridView1.Columns[7].HeaderText = "Kullanım Şekli";
            dataGridView1.Columns[8].HeaderText = "İlaç Barkod";
          
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //hastadoldur fonksiyonunu çağırır.
            hastaDoldur();
        }

        private void button8_Click(object sender, EventArgs e)
        {

            try
            {


                bag.baglan();
                DialogResult cevap;
                cevap = MessageBox.Show("Hasta Tablosunu da Boşaltılsın İstiyor Musunuz ? ", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (cevap == DialogResult.Yes)
                {

                    SqlCommand tablobosalt = new SqlCommand(" delete from hasta", bag.baglan());
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
                ffff.Graphics.DrawString("Hasta Tablosu", myFont, sbrush, 375, 65);
                ffff.Graphics.DrawLine(myPen, 50, 95, 770, 95); //çizgi çizdik.

                myFont = new System.Drawing.Font("Calibri", 6, FontStyle.Bold); //Detay başlığını yazacağımız için fontu kalın yapıp puntoyu 10 yaptık.
                ffff.Graphics.DrawString("TC Kimlik", myFont, sbrush, 50, 110); //Detay başlığı
                ffff.Graphics.DrawString("Adı Soyadı", myFont, sbrush, 125, 110); //Detay başlığı
                ffff.Graphics.DrawString("Sosyal Güvencesi", myFont, sbrush, 200, 110); // Detay başlığı
                ffff.Graphics.DrawString("Adresi", myFont, sbrush, 275, 110); //Detay başlığı
                ffff.Graphics.DrawString("Telefonu", myFont, sbrush, 425, 110); //Detay başlığı

                ffff.Graphics.DrawString("İlaç Kullanımı", myFont, sbrush, 500, 110); //Detay başlığı
                ffff.Graphics.DrawString("Kullanım Şekli", myFont, sbrush, 575, 110); //Detay başlığı
                ffff.Graphics.DrawString("İlaç Barkod", myFont, sbrush, 720, 110); //Detay başlığı
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
                    ffff.Graphics.DrawString(dataGridView1[5, i].Value.ToString(), myFont, sbrush, 425, y);//5.sütun
                    ffff.Graphics.DrawString(dataGridView1[6, i].Value.ToString(), myFont, sbrush, 500, y);//4.sütun
                    ffff.Graphics.DrawString(dataGridView1[7, i].Value.ToString(), myFont, sbrush, 575, y);//5.sütun
                    ffff.Graphics.DrawString(dataGridView1[8, i].Value.ToString(), myFont, sbrush, 720, y);//5.sütun
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
    }
}
