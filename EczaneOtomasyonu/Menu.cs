using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using Newtonsoft.Json;
using System.IO;
using System.Net;
namespace EczaneOtomasyonu
{
    public partial class Menu : Form
    {
        public Menu()
        {
            InitializeComponent();
        }
     
        private void Menu_Load(object sender, EventArgs e)
        {
            string[] jsonVerileri,bugunkiKoronaCozumle; //2 adet dizi oluşturduk

            using (WebClient wc = new WebClient())
            {
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                var json = wc.DownloadString("https://raw.githubusercontent.com/ozanerturk/covid19-turkey-api/master/dataset/timeline.json");
                //güncel verileri tutan json belgesini var değişkene aktardık
                jsonVerileri = json.ToString().Split('{');
                //{ işareti ile ayırarak bir diziye aktardık
            }
            bugunkiKoronaCozumle = jsonVerileri[jsonVerileri.Length - 1].Split('"');
            //json verisi günlük güncellendi için ve her defasında son sıradaki veride işlem yaptımız için
            //güncel veriyi verecektir.
            //son günün verisini " işareti ayırarak başka bi diziye aktardık
          
            label6.Text = bugunkiKoronaCozumle[3];
            label7.Text = bugunkiKoronaCozumle[31];
            //dizideki sıraya göre ölüm sayısı , vaka sayısı gibi sayıları bulup gerekli sıraya koyduk
            label8.Text = bugunkiKoronaCozumle[35];
            label9.Text = bugunkiKoronaCozumle[55];
            label10.Text = bugunkiKoronaCozumle[51];




        }
       
     

        private void button1_Click(object sender, EventArgs e)
        {
            ilacBilgileriForm ilacForm = new ilacBilgileriForm();
            ilacForm.Show();
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            personel_islemleri personel = new personel_islemleri();
            personel.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            hastaKayit hasta = new hastaKayit();
            hasta.Show();
            this.Hide();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            AsiTakipFormu asiTakip = new AsiTakipFormu();
            asiTakip.Show();
            this.Hide();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            hastaTakip asiTakip = new hastaTakip();
            asiTakip.Show();
            this.Hide();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            cari _cari = new cari();
            _cari.Show();
            this.Hide();
        }
    }
}
