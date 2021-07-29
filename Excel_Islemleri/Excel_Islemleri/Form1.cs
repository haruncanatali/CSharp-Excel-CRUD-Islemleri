using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel_Islemleri
{
    public partial class Form1 : Form
    {
        int id;
        public Form1()
        {
            InitializeComponent();
        }

        private void VeriTabaniBtns_Click(object sender, EventArgs e)
        {
            SimpleButton buton = sender as SimpleButton;
            switch (buton.Text)
            {
                case "Ekle":
                    Ekle_Fonk();
                    break;
                case "Güncelle":
                    Guncelle_Fonk();
                    break;
                case "Sil":
                    Sil_Fonk();
                    break;
            }
        }

        private void Sil_Fonk()
        {
            Excel_DbAccess.Sil(id);
            Yukle_Fonk();
        }

        private void Guncelle_Fonk()
        {
            Excel_DbAccess.Guncelle(new Personel
            {
                ID = id,
                AD = adTxt.Text.ToString(),
                SOYAD = soyadTxt.Text.ToString(),
                MESLEK = meslekTxt.Text.ToString(),
                MAAS = int.Parse(maasTxt.Text.ToString())
            });
            Yukle_Fonk();
        }

        private void Ekle_Fonk()
        {
            Excel_DbAccess.Ekle(new Personel
            {
                ID = (new Random()).Next(0, 10000),
                AD = adTxt.Text.ToString(),
                SOYAD = soyadTxt.Text.ToString(),
                MESLEK = meslekTxt.Text.ToString(),
                MAAS = int.Parse(maasTxt.Text.ToString())
            });
            Yukle_Fonk();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Yukle_Fonk();
        }

        private void Yukle_Fonk()
        {
            id = -1;
            adTxt.Text = "";
            soyadTxt.Text = "";
            meslekTxt.Text = "";
            maasTxt.Text = "";

            gridControl1.DataSource = Excel_DbAccess.Listele();
            
        }

        private void gridControl1_Click(object sender, EventArgs e)
        {
            try
            {
                id = int.Parse(gridView1.GetFocusedRowCellValue("ID").ToString());
                adTxt.Text = gridView1.GetFocusedRowCellValue("AD").ToString();
                soyadTxt.Text = gridView1.GetFocusedRowCellValue("SOYAD").ToString();
                meslekTxt.Text = gridView1.GetFocusedRowCellValue("MESLEK").ToString();
                maasTxt.Text = gridView1.GetFocusedRowCellValue("MAAS").ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Nesne Bulunamadı!");
            }
        }
    }
}
