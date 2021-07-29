using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Excel_Islemleri
{
    internal static class Excel_DbAccess
    {
        public static string connectionString = @"C:\Users\Acer\Desktop\VeriTabani.xlsx";
        public static List<Personel> liste;

        public  static List<Personel> Listele()
        {
            try
            {
                liste = new List<Personel>();
                FileStream fs = new FileStream(connectionString, FileMode.Open, FileAccess.Read);
                IWorkbook workbook = new XSSFWorkbook(fs);
                ISheet sheet = workbook.GetSheetAt(0);
                int satir = sheet.LastRowNum;

                for (int i = 1; i <= satir; i++)
                {
                    IRow mevcutSatir = sheet.GetRow(i);

                    liste.Add(new Personel
                    {
                        ID = int.Parse(mevcutSatir.GetCell(0).ToString().Trim()),
                        AD = mevcutSatir.GetCell(1).ToString().Trim(),
                        SOYAD = mevcutSatir.GetCell(2).ToString().Trim(),
                        MESLEK = mevcutSatir.GetCell(3).ToString().Trim(),
                        MAAS = int.Parse(mevcutSatir.GetCell(4).ToString().Trim())
                    }); 
                    
                }
                fs.Close();
                return liste;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }


        }

        public static void Ekle(Personel personel)
        {
            try
            {
                FileStream fs = new FileStream(connectionString, FileMode.Open, FileAccess.Read);
                IWorkbook workbook = new XSSFWorkbook(fs);
                ISheet sheet = workbook.GetSheetAt(0);
                int satirSayisi = sheet.LastRowNum;
                satirSayisi++;

                personel.ID = (new Random()).Next(1, 1000);

                IRow satir = sheet.CreateRow(satirSayisi);

                ICell idCell = satir.CreateCell(0);
                idCell.SetCellValue(personel.ID);

                ICell adCell = satir.CreateCell(1);
                adCell.SetCellValue(personel.AD);

                ICell soyadCell = satir.CreateCell(2);
                soyadCell.SetCellValue(personel.SOYAD);

                ICell meslekCell = satir.CreateCell(3);
                meslekCell.SetCellValue(personel.MESLEK);

                ICell maasCell = satir.CreateCell(4);
                maasCell.SetCellValue(personel.MAAS);

                using (FileStream fst = new FileStream(connectionString, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fst);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public static void Guncelle(Personel personel)
        {
            try
            {
                int index=-1;
                FileStream fs = new FileStream(connectionString, FileMode.Open, FileAccess.Read);
                IWorkbook workbook = new XSSFWorkbook(fs);
                ISheet sheet = workbook.GetSheetAt(0);
                int satir = sheet.LastRowNum;

                for (int i = 1; i <= satir; i++)
                {
                    IRow mevcutSatir = sheet.GetRow(i);

                    if(mevcutSatir.GetCell(0).ToString() == personel.ID.ToString())
                    {
                        index = i;
                        break;
                    }

                }

                IRow _satir = sheet.GetRow(index);

                ICell adCell = _satir.GetCell(1);
                ICell soyadCell = _satir.GetCell(2);
                ICell meslekCell = _satir.GetCell(3);
                ICell maasCell = _satir.GetCell(4);

                adCell.SetCellValue(personel.AD);
                soyadCell.SetCellValue(personel.SOYAD);
                meslekCell.SetCellValue(personel.MESLEK);
                maasCell.SetCellValue(personel.MAAS);

                using (FileStream fst = new FileStream(connectionString,FileMode.Create,FileAccess.Write))
                {
                    workbook.Write(fst);
                }
                
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public static void Sil(int id)
        {
            try
            {
                int index = -1;
                FileStream fs = new FileStream(connectionString, FileMode.Open, FileAccess.Read);
                IWorkbook workbook = new XSSFWorkbook(fs);
                ISheet sheet = workbook.GetSheetAt(0);
                int satir = sheet.LastRowNum;

                for (int i = 1; i <= satir; i++)
                {
                    IRow mevcutSatir = sheet.GetRow(i);

                    if (mevcutSatir.GetCell(0).ToString() == id.ToString())
                    {
                        index = i;
                        break;
                    }

                }

                IRow _satir = sheet.GetRow(index);
                sheet.RemoveRow(_satir);

                using (FileStream fst = new FileStream(connectionString, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fst);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

    }
}
