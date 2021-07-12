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
namespace EczaneOtomasyonu
{
    public partial class cari : Form
    {
        public cari()
        {
            InitializeComponent();
        }

        sqlbaglantisi bag = new sqlbaglantisi(); // sql bağlantı classımızdan nesne oluşturduk ve bağlantı için kullanacağız
        SqlCommand kmt = new SqlCommand(); //sql ekleme silme güncelleme listeleme işlemleri için sqlcommand nesnesi oluşturduk
        DataSet dtst = new DataSet();//datagridviewlere sql serverdaki tabloları aktarmak için kullanıyoruz.


        public void satilanIlacSayisi()
        {
            //satılan ilaçla beraber hasta kaydı girildi için hasta sayısı satılan ilaç sayısını verecek
            kmt.Connection = bag.baglan(); //sql komutuna bağlantı oluşturduk
            kmt.CommandText = "SELECT COUNT(*)from hasta";
            //hasta kayıt sayısını döndürecek fonksiyonu yazdık
            SqlDataReader oku;
            oku = kmt.ExecuteReader();
            if (oku.Read())
            {//hasta sayısını label6ya adeti aktardık
                label6.Text=oku[0].ToString()+" Adet";
            }

            oku.Dispose();
         

        }
        public void toplamPersonelSayisi()
        {
           
            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "SELECT COUNT(DISTINCT tc_kimlik)from personel";
            //personel sayisini tc ile tekrarlanmicak şekilde kayıt sayısını döndürecek fonksiyonu yazdık
            SqlDataReader oku;
            oku = kmt.ExecuteReader();
            if (oku.Read())
            {//buldumuz veriyi llabel10'a aktardık
                label10.Text = oku[0].ToString() + " Personel";
            }

            oku.Dispose();


        }

        public void hastaSayisiToplam()
        {
      
            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "SELECT COUNT(DISTINCT tc_kimlik)from hasta";
            //hasta sayisini tc ile tekrarlanmicak şekilde kayıt sayısını döndürecek fonksiyonu yazdık
            SqlDataReader oku;
            oku = kmt.ExecuteReader();
            if (oku.Read())
            {//buldumuz veriyi llabel9'a aktardık
                label9.Text = oku[0].ToString() + " Hasta";
            }

            oku.Dispose();


        }

        public void toplamvurulanAsi()
        {
          
            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "SELECT COUNT(*)from asiTablosu";
            //asitablosundaki kayıt sayısını döndürecek fonksiyonu yazdık
            SqlDataReader oku;
            oku = kmt.ExecuteReader();
            if (oku.Read())
            {//buldumuz veriyi label8de yazdırdık
                label8.Text = oku[0].ToString() + " Adet";
            }

            oku.Dispose();


        }
        public void kazanilanUcret()
        {
          
            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "SELECT sum(ilac.fiyati) from hasta,ilac where hasta.ilac_barkod=ilac.barkod_no";
            SqlDataReader oku;//hasta tablosundaki alınan ilaçların fiyatını ilaç tablosundan fiyatlarını çekerek toplattık
            oku = kmt.ExecuteReader();
            if (oku.Read())
            {
                //toplanan veriyi label7 ye aktardık
                label7.Text = oku[0].ToString()+" TL";
            }

            oku.Dispose();


        }
        private void cari_Load(object sender, EventArgs e)
        {//oluşturdumuz fonksiyonları form yüklenirken çağırdık
            satilanIlacSayisi();
            kazanilanUcret();
            toplamvurulanAsi();
            hastaSayisiToplam();
            toplamPersonelSayisi();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Menu menu = new Menu(); //yeni menü form oluşturduk ve açilmasini sağladık mevcut formuda kapattık
            menu.Show();
            this.Hide();
        }
    }
}
