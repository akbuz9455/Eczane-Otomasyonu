using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EczaneOtomasyonu
{
    public partial class girisForm : Form
    {
        public girisForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //oluşturmak istediğimiz kullanıcı adı ve şifreyi if koşulunun içine yazdık
            if (txtKullaniciAdi.Text == "mha24" && txtSifre.Text== "mha24")
            {
                MessageBox.Show("Giriş Başarılı !");
                Menu menu = new Menu();
                menu.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Şifre ve Kullanıcı Adı uyuşmuyor.");
            }
         
        }
    }
}
