using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;


using Word = Microsoft.Office.Interop.Word;
using System.ComponentModel;
using System.Collections.ObjectModel;

namespace IP_PPS
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.data = new PlansData();
            dic = this.data.Plans;
            this.DataContext = this.data;
        }
        PlansData data;

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            dic.Clear();
            {
                string data = File.ReadAllText("data.csv").Replace(";;;;;;;;;;;;;\r\n", "").Replace("\r\n", ";");
                var cells = data.Split(';');



                for (int i = 0; i < cells.Length - 14; i += 14)
                {
                    var name = cells[i];
                    var oplata = cells[i + 1];
                    var pred = cells[i + 6];
                    var prtype = cells[i + 7];
                    var sem = cells[i + 8];
                    var prgroups = cells[i + 9];
                    var pr_osen = cells[i + 11];
                    var pr_vesna = cells[i + 12];


                    if (!dic.ContainsKey(name))
                        dic.Add(name, new Plan());

                    dic[name].Name = name;
                    //dic[name].Trudoustr = trud;
                    dic[name].Predmets.Add(new Predmet());
                    dic[name].Predmets.Last().Name = pred;
                    dic[name].Predmets.Last().Type = prtype;
                    dic[name].Predmets.Last().Sem = ParseInt(sem);
                    dic[name].Predmets.Last().Groups = prgroups;
                    dic[name].Predmets.Last().Osen = ParseDec(pr_osen);
                    dic[name].Predmets.Last().Vesna = ParseDec(pr_vesna);
                    dic[name].Predmets.Last().Oplata = oplata;
                }


                string prdata = File.ReadAllText("prepods.csv").Replace("\r\n", ";");
                cells = prdata.Split(';');

                for (int i = 0; i < cells.Length - 1; i += 12)
                {
                    var surname = cells[i];
                    var name = cells[i + 1];
                    var stepen = cells[i + 2];
                    var dolzh = cells[i + 3];
                    var zvanie = cells[i + 4];
                    var year = ParseInt(cells[i + 5]);
                    var trud = cells[i + 6];
                    var name_rp = cells[i + 11];



                    var prepod = dic.Where(p => p.Key.Contains(surname)).Select(kv => kv.Value).FirstOrDefault();
                    if (prepod == null)
                        prepod = dic.Where(p => p.Key == name).Select(kv => kv.Value).FirstOrDefault();

                    if (prepod != null)
                    {
                        prepod.Dolzhnost = dolzh;
                        prepod.NameRP = name_rp;
                        prepod.Stepen = stepen;
                        prepod.Zvanie = zvanie;
                        if (year != 0)
                            prepod.Stazh = 2021 - year;
                        else
                            prepod.Stazh = -1;
                        prepod.Trudoustr = trud;
                        prepod.DolzhnostRP = GetDolzhnostRP(dolzh);
                    }
                }
            }
            
            this.data.OnPropertyChanged(nameof(this.data.Prepods));
            foreach (var p in this.data.Prepods)
            { 
                p.OnDistribute();
            }
        }


        

        

        Dictionary<string, Plan> dic;




        int ParseInt(string s)
        {
            int val = 0;
            int.TryParse(s, out val);
            return val;
        }
        decimal ParseDec(string s)
        {
            decimal val = 0;
            decimal.TryParse(s, out val);
            return val;
        }
        string GetDolzhnostRP(string dolzh)
        {
            dolzh = dolzh.Replace("ст.", "старшего");
            if (dolzh.Contains("доцент"))
                return "доцента";
            if (dolzh.Contains("преподаватель"))
                return dolzh.Replace('ь', 'я');
            if (dolzh.Contains("ассистент"))
                return "ассистента";
            if (dolzh.Contains("зав"))
                return "заведующего кафедрой";
            return dolzh + 'а';
        }


        
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            data.SelectedPlan.Foses.Add(new Plan.Rabota(data.SelectedPlan.UpdateFoses));
        }
        private void RButton_Click_1(object sender, RoutedEventArgs e)
        {
            if(data.SelectedPlan.SelectedFos != null)
                data.SelectedPlan.Foses.Remove(data.SelectedPlan.SelectedFos);
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            data.SelectedPlan.NauchOrg.Add(new Plan.Rabota(data.SelectedPlan.UpdateNauchOrg));
        }
        private void RButton_Click_2(object sender, RoutedEventArgs e)
        {
            if (data.SelectedPlan.SelectedNauchOrg != null)
                data.SelectedPlan.NauchOrg.Remove(data.SelectedPlan.SelectedNauchOrg);
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            data.SelectedPlan.NauchIssl.Add(new Plan.Rabota(data.SelectedPlan.UpdateNauchIssl));
        }
        private void RButton_Click_3(object sender, RoutedEventArgs e)
        {
            if (data.SelectedPlan.SelectedNauchIssl != null)
                data.SelectedPlan.NauchIssl.Remove(data.SelectedPlan.SelectedNauchIssl);
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            data.SelectedPlan.NauchMetod.Add(new Plan.Rabota(data.SelectedPlan.UpdateNauchMetod));
        }
        private void RButton_Click_4(object sender, RoutedEventArgs e)
        {
            if (data.SelectedPlan.SelectedNauchMetod != null)
                data.SelectedPlan.NauchMetod.Remove(data.SelectedPlan.SelectedNauchMetod);
        }

        private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            if(comboBox.SelectedIndex > 0)
                comboBox.SelectedIndex--;
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            if (comboBox.SelectedIndex < comboBox.Items.Count - 1)
                comboBox.SelectedIndex++;
        }




        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            data.SelectedPlan.Ispob.Add(new Plan.Rabota(data.SelectedPlan.UpdateIspob));
        }
        private void RButton_Click_7(object sender, RoutedEventArgs e)
        {
            if (data.SelectedPlan.SelectedIspob != null)
                data.SelectedPlan.Ispob.Remove(data.SelectedPlan.SelectedIspob);
        }
        private void Button_Click_8(object sender, RoutedEventArgs e)
        {
            data.SelectedPlan.Kval.Add(new Plan.Rabota(data.SelectedPlan.UpdateKval));
        }
        private void RButton_Click_8(object sender, RoutedEventArgs e)
        {
            if (data.SelectedPlan.SelectedKval != null)
                data.SelectedPlan.Kval.Remove(data.SelectedPlan.SelectedKval);
        }



       
        private void Button_Click_9(object sender, RoutedEventArgs e)
        {
            Word.Application app = new Microsoft.Office.Interop.Word.Application { Visible = true };
            Word.Document template = app.Documents.Open(
                System.AppDomain.CurrentDomain.BaseDirectory + @"\Document.docx",
                ReadOnly: false, Visible: true);
            template.Activate();


            var plan = this.data.SelectedPlan;


            template.SaveAs2(System.AppDomain.CurrentDomain.BaseDirectory + $@"\Планы\ИТИВС 2021 Индивидуальный план {plan.NameRP}.docx");







            Word.Document doc = app.Documents.Open(
                System.AppDomain.CurrentDomain.BaseDirectory + $@"\Планы\ИТИВС 2021 Индивидуальный план {plan.NameRP}.docx",
                ReadOnly: false, Visible: true);
            doc.Activate();
            doc.SelectAllEditableRanges();


            app.Replace("dolzhnost", plan.Dolzhnost);
            app.Replace("dolzhrp", plan.DolzhnostRP);
            app.Replace("fiorp", plan.NameRP);
            app.Replace("zvanie", plan.Zvanie);
            app.Replace("stepen", plan.Stepen);
            app.Replace("stazh", plan.Stazh);
            app.Replace("stavka", plan.Stavka);
            app.Replace("trudoustr", plan.Trudoustr);



            if (plan.GroupedPredmets.Count > 20)
                System.Windows.MessageBox.Show("НУЖНО БОЛЬШЕ ЧЕМ 20 ЗАГЛУШЕК В ШАБЛОНЕ ДЛЯ ТАБЛИЦЫ 1.2");
            for (int pr = 1, i = 0; pr <= 20; pr++, i++)
            {
                var gps = plan.GroupedPredmets;
                if (i < gps.Count)
                {
                    var gp = gps[i];
                    app.Replace($"{pr}pr_sem", gp.Sem);
                    app.Replace($"{pr}pr_nazv",
                        $"{gp.Name} {string.Join(", ", gp.Groups.Select(group => $"({group})").ToArray())}");
                    for (int k = 0; k <= 12; k++)
                    {
                        string str;
                        if (gp.Hours[k] == 0)
                            str = "-";
                        else
                            str = gp.Hours[k].ToString();
                        app.Replace($"{pr}pr_{k + 1}", str);
                    }
                }
                else
                    break;
                //else
                //{
                //    app.Replace($"{pr}pr_sem", "");
                //    app.Replace($"{pr}pr_nazv", "");
                //    for (int k = 0; k <= 12; k++)
                //    {
                //        app.Replace($"{pr}pr_{k + 1}", "");
                //    }
                //}
            }
            var tables = doc.Tables;
            var tb = tables[3];
            var rowstodelete = 20 - plan.GroupedPredmets.Count;
            for (int i = 0; i < rowstodelete; i++)
            { 
                tb.Cell(plan.GroupedPredmets.Count + 3, 1).Select();
                app.Selection.SelectRow();
                app.Selection.Cells.Delete();
            }

            {
                app.Replace("lek_osen", plan.GroupedPredmets
                .Where(p => p.Period == "осень")
                .Select(p => p.Hours[0])
                .Sum().ToHours());
                app.Replace("lab_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[1])
                    .Sum().ToHours());
                app.Replace("sem_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[2])
                    .Sum().ToHours());
                app.Replace("kurs_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[3])
                    .Sum().ToHours());
                app.Replace("ind_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[4])
                    .Sum().ToHours());
                app.Replace("prakt_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[5])
                    .Sum().ToHours());
                app.Replace("pred_spec_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[6])
                    .Sum().ToHours());
                app.Replace("vkr_spec_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[7])
                    .Sum().ToHours());
                app.Replace("pred_bak_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[8])
                    .Sum().ToHours());
                app.Replace("vkr_bak_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[9])
                    .Sum().ToHours());
                app.Replace("mag_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[10])
                    .Sum().ToHours());
                app.Replace("mag_kons_osen", plan.GroupedPredmets
                   .Where(p => p.Period == "осень")
                   .Select(p => p.Hours[11])
                   .Sum().ToHours());
                app.Replace("vkr_mag_osen", plan.GroupedPredmets
                   .Where(p => p.Period == "осень")
                   .Select(p => p.Hours[12])
                   .Sum().ToHours());

                app.Replace("uch_vsego_osen", plan.GroupedPredmets
                   .Where(p => p.Period == "осень")
                   .Select(p => p.Hours.Sum())
                   .Sum().ToHours());
            }
            
            {
                app.Replace("lek_vesna", plan.GroupedPredmets
                .Where(p => p.Period == "весна")
                .Select(p => p.Hours[0])
                .Sum().ToHours());
                app.Replace("lab_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[1])
                    .Sum().ToHours());
                app.Replace("sem_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[2])
                    .Sum().ToHours());
                app.Replace("kurs_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[3])
                    .Sum().ToHours());
                app.Replace("ind_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[4])
                    .Sum().ToHours());
                app.Replace("prakt_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[5])
                    .Sum().ToHours());
                app.Replace("pred_spec_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[6])
                    .Sum().ToHours());
                app.Replace("vkr_spec_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[7])
                    .Sum().ToHours());
                app.Replace("pred_bak_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[8])
                    .Sum().ToHours());
                app.Replace("vkr_bak_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[9])
                    .Sum().ToHours());
                app.Replace("mag_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[10])
                    .Sum().ToHours());
                app.Replace("mag_kons_vesna", plan.GroupedPredmets
                   .Where(p => p.Period == "весна")
                   .Select(p => p.Hours[11])
                   .Sum().ToHours());
                app.Replace("vkr_mag_vesna", plan.GroupedPredmets
                   .Where(p => p.Period == "весна")
                   .Select(p => p.Hours[12])
                   .Sum().ToHours());

                app.Replace("uch_vsego_vesna", plan.GroupedPredmets
                   .Where(p => p.Period == "весна")
                   .Select(p => p.Hours.Sum())
                   .Sum().ToHours());
            }

            {
                app.Replace("lek_vsego", plan.GroupedPredmets

                .Select(p => p.Hours[0])
                .Sum().ToHours());
                app.Replace("lab_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[1])
                    .Sum().ToHours());
                app.Replace("sem_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[2])
                    .Sum().ToHours());
                app.Replace("kurs_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[3])
                    .Sum().ToHours());
                app.Replace("ind_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[4])
                    .Sum().ToHours());
                app.Replace("prakt_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[5])
                    .Sum().ToHours());
                app.Replace("pred_spec_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[6])
                    .Sum().ToHours());
                app.Replace("vkr_spec_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[7])
                    .Sum().ToHours());
                app.Replace("pred_bak_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[8])
                    .Sum().ToHours());
                app.Replace("vkr_bak_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[9])
                    .Sum().ToHours());
                app.Replace("mag_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[10])
                    .Sum().ToHours());
                app.Replace("mag_kons_vsego", plan.GroupedPredmets

                   .Select(p => p.Hours[11])
                   .Sum().ToHours());
                app.Replace("vkr_mag_vsego", plan.GroupedPredmets

                   .Select(p => p.Hours[12])
                   .Sum().ToHours());

                app.Replace("uch_vsego", plan.GroupedPredmets

                   .Select(p => p.Hours.Sum())
                   .Sum().ToHours());
            }

            if (plan.GroupedPredmetsDop.Count > 20)
                System.Windows.MessageBox.Show("НУЖНО БОЛЬШЕ ЧЕМ 20 ЗАГЛУШЕК В ШАБЛОНЕ ДЛЯ ТАБЛИЦЫ 2.2");
            for (int pr = 1, i = 0; pr <= 20; pr++, i++)
            {
                var gps = plan.GroupedPredmetsDop;
                if (i < gps.Count)
                {
                    var gp = gps[i];
                    app.Replace($"{pr}pd_sem", gp.Sem);
                    app.Replace($"{pr}pd_nazv",
                        $"{gp.Name} {string.Join(", ", gp.Groups.Select(group => $"({group})").ToArray())}");
                    for (int k = 0; k <= 12; k++)
                    {
                        string str;
                        if (gp.Hours[k] == 0)
                            str = "-";
                        else
                            str = gp.Hours[k].ToString();
                        app.Replace($"{pr}pd_{k + 1}", str);
                    }
                }
                else
                    break;
                //else
                //{
                //    app.Replace($"{pr}pd_sem", "");
                //    app.Replace($"{pr}pd_nazv", "");
                //    for (int k = 0; k <= 12; k++)
                //    {
                //        app.Replace($"{pr}pd_{k + 1}", "");
                //    }
                //}
            }
            tables = doc.Tables;
            tb = tables[5];
            rowstodelete = 20 - plan.GroupedPredmetsDop.Count;
            for (int i = 0; i < rowstodelete; i++)
            {
                tb.Cell(plan.GroupedPredmetsDop.Count + 3, 1).Select();
                app.Selection.SelectRow();
                app.Selection.Cells.Delete();
            }

            {
                app.Replace("dlek_osen", plan.GroupedPredmetsDop
                .Where(p => p.Period == "осень")
                .Select(p => p.Hours[0])
                .Sum().ToHours());
                app.Replace("dlab_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[1])
                    .Sum().ToHours());
                app.Replace("dsem_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[2])
                    .Sum().ToHours());
                app.Replace("dkurs_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[3])
                    .Sum().ToHours());
                app.Replace("dind_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[4])
                    .Sum().ToHours());
                app.Replace("dprakt_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[5])
                    .Sum().ToHours());
                app.Replace("dpred_spec_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[6])
                    .Sum().ToHours());
                app.Replace("dvkr_spec_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[7])
                    .Sum().ToHours());
                app.Replace("dpred_bak_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[8])
                    .Sum().ToHours());
                app.Replace("dvkr_bak_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[9])
                    .Sum().ToHours());
                app.Replace("dmag_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[10])
                    .Sum().ToHours());
                app.Replace("dmag_kons_osen", plan.GroupedPredmetsDop
                   .Where(p => p.Period == "осень")
                   .Select(p => p.Hours[11])
                   .Sum().ToHours());
                app.Replace("dvkr_mag_osen", plan.GroupedPredmetsDop
                   .Where(p => p.Period == "осень")
                   .Select(p => p.Hours[12])
                   .Sum().ToHours());

                app.Replace("duch_vsego_osen", plan.GroupedPredmetsDop
                   .Where(p => p.Period == "осень")
                   .Select(p => p.Hours.Sum())
                   .Sum().ToHours());
            }

            {
                app.Replace("dlek_vesna", plan.GroupedPredmetsDop
                .Where(p => p.Period == "весна")
                .Select(p => p.Hours[0])
                .Sum().ToHours());
                app.Replace("dlab_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[1])
                    .Sum().ToHours());
                app.Replace("dsem_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[2])
                    .Sum().ToHours());
                app.Replace("dkurs_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[3])
                    .Sum().ToHours());
                app.Replace("dind_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[4])
                    .Sum().ToHours());
                app.Replace("dprakt_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[5])
                    .Sum().ToHours());
                app.Replace("dpred_spec_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[6])
                    .Sum().ToHours());
                app.Replace("dvkr_spec_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[7])
                    .Sum().ToHours());
                app.Replace("dpred_bak_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[8])
                    .Sum().ToHours());
                app.Replace("dvkr_bak_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[9])
                    .Sum().ToHours());
                app.Replace("dmag_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[10])
                    .Sum().ToHours());
                app.Replace("dmag_kons_vesna", plan.GroupedPredmetsDop
                   .Where(p => p.Period == "весна")
                   .Select(p => p.Hours[11])
                   .Sum().ToHours());
                app.Replace("dvkr_mag_vesna", plan.GroupedPredmetsDop
                   .Where(p => p.Period == "весна")
                   .Select(p => p.Hours[12])
                   .Sum().ToHours());

                app.Replace("duch_vsego_vesna", plan.GroupedPredmetsDop
                   .Where(p => p.Period == "весна")
                   .Select(p => p.Hours.Sum())
                   .Sum().ToHours());
            }

            {
                app.Replace("dlek_vsego", plan.GroupedPredmetsDop

                .Select(p => p.Hours[0])
                .Sum().ToHours());
                app.Replace("dlab_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[1])
                    .Sum().ToHours());
                app.Replace("dsem_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[2])
                    .Sum().ToHours());
                app.Replace("dkurs_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[3])
                    .Sum().ToHours());
                app.Replace("dind_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[4])
                    .Sum().ToHours());
                app.Replace("dprakt_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[5])
                    .Sum().ToHours());
                app.Replace("dpred_spec_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[6])
                    .Sum().ToHours());
                app.Replace("dvkr_spec_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[7])
                    .Sum().ToHours());
                app.Replace("dpred_bak_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[8])
                    .Sum().ToHours());
                app.Replace("dvkr_bak_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[9])
                    .Sum().ToHours());
                app.Replace("dmag_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[10])
                    .Sum().ToHours());
                app.Replace("dmag_kons_vsego", plan.GroupedPredmetsDop

                   .Select(p => p.Hours[11])
                   .Sum().ToHours());
                app.Replace("dvkr_mag_vsego", plan.GroupedPredmetsDop

                   .Select(p => p.Hours[12])
                   .Sum().ToHours());

                app.Replace("duch_vsego", plan.GroupedPredmetsDop

                   .Select(p => p.Hours.Sum())
                   .Sum().ToHours());
            }


            for (int pr = 1, i = 0; pr <= 8; pr++, i++)
            {
                var mps = plan.MetodPredmets;
                if (i < mps.Count)
                {
                    var mp = mps[i];
                    app.Replace($"metod_pr{pr}_nazv", mp.Name);
                    app.Replace($"metod_pr{pr}_osen", mp.Osen.ToHours());
                    app.Replace($"metod_pr{pr}_vesna", mp.Vesna.ToHours());

                    app.Replace($"metod_pr{pr}_osen_per", "осень");
                    app.Replace($"metod_pr{pr}_vesna_per", "весна");
                }
                else
                {
                    app.Replace($"metod_pr{pr}_nazv", "");
                    app.Replace($"metod_pr{pr}_osen", "");
                    app.Replace($"metod_pr{pr}_vesna", "");

                    app.Replace($"metod_pr{pr}_osen_per", "");
                    app.Replace($"metod_pr{pr}_vesna_per", "");
                }
            }
            for (int pr = 1, i = 0; pr <= 5; pr++, i++)
            {
                var mps = plan.foses;
                if (i < mps.Count)
                {
                    var mp = mps[i];
                    app.Replace($"fosrpd{pr}_nazv", mp.Fos);
                    app.Replace($"fosrpd{pr}", mp.Hours);
                    app.Replace($"fosrpd{pr}per", mp.Period);

                }
                else
                {
                    app.Replace($"fosrpd{pr}_nazv", "");
                    app.Replace($"fosrpd{pr}", "");
                    app.Replace($"fosrpd{pr}per", "");
                }
            }
            app.Replace("4vsego_osen", plan.foses
                .Where(f => f.Period == "осень")
                .Select(f => f.Hours).Sum()
                +
                plan.MetodPredmets
                .Select(p => p.Osen).Sum()
                );
            app.Replace("4vsego_vesna", plan.foses
                .Where(f => f.Period == "весна")
                .Select(f => f.Hours).Sum()
                +
                plan.MetodPredmets
                .Select(p => p.Vesna).Sum()
                );
            app.Replace("4vsego", plan.foses
                .Select(f => f.Hours).Sum()
                +
                plan.MetodPredmets
                .Select(p => p.Osen + p.Vesna).Sum()
                );




            doc.Save();
            System.Windows.MessageBox.Show($"План {plan.NameRP} сгенерирован!");
        }
    }


    class Base : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    public static class DocExtensions
    {
        private static void FindAndReplace(Microsoft.Office.Interop.Word.Application doc, object findText, object replaceWithText)
        {
            //options
            object matchCase = true;
            object matchWholeWord = false;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            doc.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }

        public static void Replace(this Word.Application doc, string key, object value)
        {
            FindAndReplace(doc, $"%{key}%", value.ToString());
        }
        public static void Replace(this Word.Application doc, string key, string value)
        {
            FindAndReplace(doc, $"%{key}%", value);
        }


        public static void ReplaceStr(this Word.Application doc, string str, string value)
        {
            FindAndReplace(doc, str, value);
        }


        public static string ToHours(this decimal hours)
        {
            if (hours == 0)
                return "-";
            else
                return hours.ToString();
        }
        public static decimal Normalize(this decimal value)
        {
            return value / 1.000000000000000000000000000000000m;
        }
    }
}
