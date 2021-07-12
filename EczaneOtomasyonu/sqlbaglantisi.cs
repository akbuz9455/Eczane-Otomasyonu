using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
namespace EczaneOtomasyonu
{
    class sqlbaglantisi
    {
        public SqlConnection baglan()
        {
            SqlConnection baglanti = new SqlConnection("Data Source=.; initial Catalog=eczane; Integrated Security=true");
            //sql bağlantı komutumuzu oluşturduk
            baglanti.Open();//bağlantıyı açtık
            SqlConnection.ClearPool(baglanti);
            SqlConnection.ClearAllPools();
            //geçmiş bağlantıları temizledik
            return (baglanti);//bağlantıyı döndürdük
        }
    }
}
