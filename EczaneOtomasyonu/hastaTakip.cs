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
    public partial class hastaTakip : Form
    {
        public hastaTakip()
        {
            InitializeComponent();
        }
        sqlbaglantisi bag = new sqlbaglantisi(); // sql bağlantı classımızdan nesne oluşturduk ve bağlantı için kullanacağız
        SqlCommand kmt = new SqlCommand(); //sql ekleme silme güncelleme listeleme işlemleri için sqlcommand nesnesi oluşturduk
        DataSet dtst = new DataSet();//datagridviewlere sql serverdaki tabloları aktarmak için kullanıyoruz.
        DataSet dtst2 = new DataSet();//3 ayrı tabloda işlem yapabilcem için fazladan 2 tane daha oluşturduk
        DataSet dtst3 = new DataSet();
        public void hastaTC()
        {

            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "SELECT DISTINCT * from hasta";//tekrarlanmicak şekilde hastaları listeledik
            SqlDataReader oku;
            oku = kmt.ExecuteReader();
            while (oku.Read())
            {//bulunan tcleri combobox1e aktardık
                comboBox1.Items.Add(oku[1].ToString());
            }

            oku.Dispose();
            comboBox1.Sorted = true;

        }
        private void button3_Click(object sender, EventArgs e)
        {
            Menu menu = new Menu(); //yeni menü form oluşturduk ve açilmasini sağladık mevcut formuda kapattık
            menu.Show();
            this.Hide();
        }

        private void hastaTakip_Load(object sender, EventArgs e)
        {//oluşturdumuz fonksiyonları form yüklenirken çağırdık
            hastaTC();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "SELECT DISTINCT * from hasta where tc_kimlik ='" + comboBox1.Text + "'";
            //tekrarlanmicak şekilde seçili comboboxdaki tcye göre listeleme komutu yazdık
            SqlDataReader oku;
            oku = kmt.ExecuteReader();
            if (oku.Read())//buldu kayıttaki adsoyad ve adres bilgilerini gerekli yerlere yazdırdık
            {
                label3.Text = oku[2].ToString();
                richTextBox1.Text = oku[4].ToString();
            }

            oku.Dispose();
        }

        private void button5_Click(object sender, EventArgs e) //vurulan aşıları listeliyoruz yani aşı vutan tc sini asi tablosunda arıyoruz.
        {
            dtst.Clear();
            //dataseti temizledik
            SqlDataAdapter adtr = new SqlDataAdapter("select asiAdi,etkiSuresi,etkisi,asiVurulmaTarihi From asiTablosu where vurulanTC='" + comboBox1.Text + "'", bag.baglan());
            adtr.Fill(dtst, "asiTablosu");
            dataGridView1.DataMember = "asiTablosu";
            dataGridView1.DataSource = dtst;
            adtr.Dispose(); //ardından yazdımız arama komutuna göre datagridwiewi doldurduk
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            //seçili satırı tamamen seçmesini ayarladık ve alttaki komutlarımızda kolon başlıklarını düzenledik
            dataGridView1.BackgroundColor = Color.White;
            dataGridView1.RowHeadersVisible = false;
            
          
            dataGridView1.Columns[0].HeaderText = "Aşı Adı";
            dataGridView1.Columns[1].HeaderText = "Etki Süresi";
            dataGridView1.Columns[2].HeaderText = "Etkisi";
            dataGridView1.Columns[3].HeaderText = "Aşı vurulma Tarihi";
        }

        private void button1_Click(object sender, EventArgs e) // hastanın aldığı ilaçları listeledik
        {
            dtst2.Clear();
            //dataseti temizledik
            SqlDataAdapter adtr = new SqlDataAdapter("select ilac.ilacin_adi From hasta,ilac where hasta.tc_kimlik='" + comboBox1.Text + "' and hasta.ilac_barkod=ilac.barkod_no", bag.baglan());

            //buradaki sql komutumuzda 2 tabloyu birleştirdik.ve hasta tablosundan yazılan barkod noyu alıp ilaç tablosundan bunun bilgisini çektik
            adtr.Fill(dtst2, "asiTablosu");
            dataGridView1.DataMember = "asiTablosu";
            dataGridView1.DataSource = dtst2;
            adtr.Dispose(); //ardından yazdımız arama komutuna göre datagridwiewi doldurduk
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            //seçili satırı tamamen seçmesini ayarladık ve alttaki komutlarımızda kolon başlıklarını düzenledik
            dataGridView1.BackgroundColor = Color.White;
            dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dataGridView1.Columns[0].HeaderText = "Hastalığa Göre Verilen İlaç ";
        }

        private void button2_Click(object sender, EventArgs e) //hasta geçirdiği hastalıkları listeleme fonksiyonumuz
        {

            dtst3.Clear();

            SqlDataAdapter adtr = new SqlDataAdapter("select ilac.kullanim_amaci From hasta,ilac where hasta.tc_kimlik='" + comboBox1.Text + "' and hasta.ilac_barkod=ilac.barkod_no", bag.baglan());
            //buradaki sql komutumuzda 2 tabloyu birleştirdik.ve hasta tablosundan yazılan barkod noyu alıp ilaç tablosundan bunun bilgisini çektik ve ilaç etkisini çekerek hastalığı öğrendik
            adtr.Fill(dtst3, "asiTablosu");
            dataGridView1.DataMember = "asiTablosu";
            dataGridView1.DataSource = dtst3;
            adtr.Dispose();//ardından yazdımız arama komutuna göre datagridwiewi doldurduk
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            //seçili satırı tamamen seçmesini ayarladık ve alttaki komutlarımızda kolon başlıklarını düzenledik
            dataGridView1.BackgroundColor = Color.White;
            dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dataGridView1.Columns[0].HeaderText = "Hastalığa Göre Verilen İlaç Etkisi";
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {


                bag.baglan();
              
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
    }
}
