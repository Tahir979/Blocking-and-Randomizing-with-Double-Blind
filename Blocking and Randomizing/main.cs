using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Collections;
using System.Diagnostics;

namespace Blocking_and_Randomizing
{
    public partial class main : MetroFramework.Forms.MetroForm
    {
        public main()
        {
            InitializeComponent();
        }

        string filename;
        int count, b = 0, y = 0, mod = 0, e = 0, t = 0, diger = 0, u = 0, p = 0, koru = 0;

        void KillSpecificExcelFileProcess(string excelFileName)
        {
            var processes = from p in Process.GetProcessesByName("EXCEL") select p;

            foreach (var process in processes)
            {
                if (process.MainWindowTitle == "")
                    process.Kill();
            }
        }
        void move()
        {
            if (y == 0 && mtgridparticipant.Rows.Count != 0)
            {
                t++;

                if (t == 1 || t == 2)
                {
                    int v = mtgridparticipant.Rows.Count;

                    mtgridparticipant.Rows.RemoveAt(v - 1);
                    mtgridparticipant.Refresh();

                    if (t == 2)
                    {
                        y++;
                    }
                }
            }
        }
        void atama()
        {
            //giriş
            #region
            katilimci_tamami.Items.Clear();

            //demek listin dinamik değişken olarak kullanımı da varmış, ilginç
            List<string> birinci = new List<string>();
            List<string> ikinci = new List<string>();
            List<string> ucuncu = new List<string>();
            List<string> dorduncu = new List<string>();
            List<int> toplam = new List<int>();

            int[] dizi = new int[mtgridparticipant.Rows.Count];

            for (int i = 0; i < mtgridparticipant.Rows.Count; i++)
            {
                if (mtgridparticipant.Rows[i].Cells[2].Value.ToString() == "")
                {

                }
                else
                {
                    dizi[i] = Convert.ToInt32(mtgridparticipant.Rows[i].Cells[2].Value.ToString());
                }
            }
            #endregion

            //dizideki sayıları sıralama
            #region
            for (int i = 0; i < dizi.Length - 1; i++)
            {
                for (int j = i; j < dizi.Length; j++)
                {
                    if (dizi[i] >= dizi[j])
                    {
                        int gecici = dizi[j];
                        dizi[j] = dizi[i];
                        dizi[i] = gecici;
                    }
                }
            }

            for (int i = 0; i < dizi.Length; i++)
            {
                katilimci_tamami.Items.Add(dizi[i]);
            }
            #endregion

            //grup sayılarına göre veritabanının ayarlanması
            #region
            if (metroTextBox3.Text.Contains(';') == true)
            {
                for (int i = 0; i < metroTextBox3.Text.Length; i++)
                {
                    if (metroTextBox3.Text[i].ToString() == ";")
                    {
                        count++;
                    }
                }
            }

            mtgridgroups.ColumnCount = Convert.ToInt32(metroTextBox1.Text);

            for (int m = 0; m < Convert.ToInt32(metroTextBox1.Text); m++)
            {
                mtgridgroups.Columns[m].HeaderText = "Grup " + Convert.ToInt32(m + 1);
            }
            #endregion

            //gruplara atama
            #region
            for (int i = 0; i <= count; i++)
            {
                if (metroTextBox3.Text.Contains(',') == true)
                {
                    //bu değişkenlerin denemesini yaptım hepsi gayet düzgün bir biçimde çalışıyor şu an
                    birinci.Add(metroTextBox3.Text.Split(';')[i]); //bu parantez içindeki 0'lar işte bizim değişkenlerimiz olacak giderek artacak, her değişkenden sonra da silinmesi gerekecek list'in
                    ikinci.Add(birinci[0].Split(',')[0]);
                    ucuncu.Add(birinci[0].Substring(ikinci[0].Length, birinci[0].Length - ikinci[0].Length));
                    dorduncu.Add(ucuncu[0].Substring(1));

                    for (int t = 0; t < katilimci_tamami.Items.Count; t++)
                    {
                        if (Convert.ToInt32(katilimci_tamami.Items[t].ToString()) <= Convert.ToInt32(dorduncu[0].ToString()) && Convert.ToInt32(katilimci_tamami.Items[t].ToString()) >= Convert.ToInt32(ikinci[0].ToString()))
                        {
                            toplam.Add(Convert.ToInt32(katilimci_tamami.Items[t]));
                        }
                    }

                    for (int b = 0; b < toplam.Count; b++)
                    {
                        katilimci_kosul.Items.Add(toplam[b]);
                    }

                    ArrayList list = new ArrayList();
                    foreach (object o in katilimci_kosul.Items)
                    {
                        list.Add(o);
                    }
                    list.Sort();
                    katilimci_kosul.Items.Clear();
                    foreach (object o in list)
                    {
                        katilimci_kosul.Items.Add(o);
                    }

                    //kümülatif ilerlemesi lazım listbox değerlerine eklemenin
                    //ilkinden kalan 1 hala devam ediyor sonuç olarak ve onun da grup sayısı ile olan moduna bakılması lazım ki artan cell sayısı var mı bilmemiz lazım
                    mod = katilimci_kosul.Items.Count % Convert.ToInt32(metroTextBox1.Text);
                    listBox1.Items.Add(mod);
                    int[] dizi_mod = new int[listBox1.Items.Count];

                    for (int z = 0; z < listBox1.Items.Count; z++)
                    {
                        dizi_mod[z] = Convert.ToInt32(listBox1.Items[z]);
                    }

                    int total = dizi_mod.Sum();
                    lst_toplam.Items.Add(total);

                    if (total >= Convert.ToInt32(metroTextBox1.Text))
                    {
                        int total_mod = total % Convert.ToInt32(metroTextBox1.Text);
                        int f = lst_toplam.Items.Count;

                        for (int g = 0; g <= lst_toplam.Items.Count; g++)
                        {
                            if (g == f)
                            {
                                lst_toplam.Items[g - 1] = total_mod;
                            }
                        }
                    }

                    if (mod == 0)
                    {
                        for (int j = 0; j < katilimci_kosul.Items.Count / Convert.ToInt32(metroTextBox1.Text); j++)
                        {
                            mtgridgroups.Rows.Add("", "");
                        }
                    }
                    else if (mod != 0)
                    {
                        for (int j = 0; j < katilimci_kosul.Items.Count / Convert.ToInt32(metroTextBox1.Text); j++)
                        {
                            mtgridgroups.Rows.Add("", "");
                        }

                        mtgridgroups.Rows.Add("", "");
                    }

                    int arttir = 0;
                    //gruplara puanlar atandı ve puanlar bloklandı
                    for (int r = 0; r < katilimci_kosul.Items.Count - 1; r++)
                    {
                        if (r == 0)
                        {

                        }

                        else
                        {
                            r++;
                        }

                        if (diger != 0)
                        {
                            if (koru == Convert.ToInt32(lst_toplam.Items[diger - 1]) && p != 0) // burayı sadece son satır durumuna göre özelleştirmem lazım
                            {//YESSSS, BAŞARDIM AQ
                                arttir++;
                                if (arttir == 2)
                                {
                                    r = 1;
                                    p = 0;
                                }
                            }
                        }

                        for (e = 0; e < Convert.ToInt32(metroTextBox1.Text); e++)
                        {
                            // eğer r katilimci kosuldan büyük olursa eklemesin

                            if (r == katilimci_kosul.Items.Count)
                            {
                                p = 0;
                                break;
                            }
                            else
                            {
                                if (diger != 0 && u == 0)
                                {
                                    e += Convert.ToInt32(lst_toplam.Items[diger - 1]);
                                    u++;
                                }

                                mtgridgroups.Rows[b].Cells[e].Value = katilimci_kosul.Items[r].ToString();
                                r++;

                                if (e == Convert.ToInt32(metroTextBox1.Text) - 1)
                                {
                                    b++;
                                }
                            }
                        }

                        r -= 2;
                    }

                    birinci.Clear();
                    ikinci.Clear();
                    ucuncu.Clear();
                    dorduncu.Clear();
                    toplam.Clear();
                    katilimci_kosul.Items.Clear();
                    diger++;
                    u = 0;
                    p++;
                    Array.Clear(dizi_mod, 0, dizi_mod.Length);
                    koru = Convert.ToInt32(metroTextBox1.Text) - 1;
                }
            }
            #endregion
        }

        private void MetroTile1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openfile1 = new OpenFileDialog
                {
                    Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls",
                    Title = "Veri Excel'ini seçiniz..."
                };

                if (openfile1.ShowDialog() == DialogResult.OK)
                {
                    filename = openfile1.FileName;
                }

                Excel.Application oXL = new Excel.Application();
                if (filename == string.Empty)
                {
                    return;
                }

                Excel.Workbook oWB = oXL.Workbooks.Open(filename);

                List<string> liste = new List<string>();
                foreach (Excel.Worksheet oSheet in oWB.Worksheets)
                {
                    liste.Add(oSheet.Name);
                }
                oWB.Close();
                oXL.Quit();
                oWB = null;
                oXL = null;
                metroGrid1.DataSource = liste.Select(x => new { SayfaAdi = x }).ToList();

                OleDbCommand komut = new OleDbCommand();
                string pathconn = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + filename + ";Extended Properties=\"Excel 8.0;HDR= yes;\";";
                OleDbConnection conn = new OleDbConnection(pathconn);
                OleDbDataAdapter MyDataAdapter = new OleDbDataAdapter("Select * from [" + metroGrid1.Rows[0].Cells[0].Value.ToString() + "$]", conn);
                DataTable dt3 = new DataTable();
                MyDataAdapter.Fill(dt3);
                mtgridparticipant.DataSource = dt3;

                for (int i = 0; i < mtgridparticipant.Rows.Count; i++)
                {
                    if (string.IsNullOrEmpty(mtgridparticipant.Rows[i].Cells[0].Value.ToString()) == true)
                    {
                        mtgridparticipant.Rows.RemoveAt(i);
                        mtgridparticipant.Refresh();
                    }
                }

                KillSpecificExcelFileProcess(filename);
            }

            catch (Exception)
            {
                return;
            }
        }
        private void metroButton3_Click(object sender, EventArgs e)
        {
            atama();
        }
        private void metroButton3_MouseMove(object sender, MouseEventArgs e)
        {
            move();
        }
        private void main_MouseMove(object sender, MouseEventArgs e)
        {
            move();
        }
        private void mtgridparticipant_MouseMove(object sender, MouseEventArgs e)
        {
            move();
        }
        private void mtgridgroups_MouseMove(object sender, MouseEventArgs e)
        {
            move();
        }
        private void MetroTile1_MouseMove(object sender, MouseEventArgs e)
        {
            move();
        }
        private void mtgridparticipant_ColumnAdded(object sender, DataGridViewColumnEventArgs e)
        {
            mtgridparticipant.Columns[e.Column.Index].SortMode = DataGridViewColumnSortMode.NotSortable;
        }
    }
}
