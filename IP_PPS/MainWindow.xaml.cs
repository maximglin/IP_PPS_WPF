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
        }


        class Predmet
        {
            public string Name;
            public string Groups;
            public string Type;
            public int Sem;
            public decimal Osen;
            public decimal Vesna;

            public string Oplata;

            public decimal Vsego => Osen + Vesna;
        }

        class GroupedPredmet
        {
            public string Name { get; set; }
            public HashSet<string> Groups { get; } = new HashSet<string>();
            public int Sem { get; set; }
            public string Period => Sem % 2 == 1 ? "осень" : "весна";
            public decimal[] Hours { get; } = new decimal[13];
        }
        class Plan : Base
        {
            public string Name { get; set; } = string.Empty;

            public List<Predmet> Predmets { get; } = new List<Predmet>();

            public decimal Stavka => Hours / 900m;
            public decimal Hours => PredmetsShtat.Select(p => p.Vsego).Sum();
            public decimal HoursToCount => Stavka * 600m;

            public List<Predmet> PredmetsShtat => Predmets.Where(p => !p.Oplata.Contains("стимулир")).ToList();
            public List<Predmet> PredmetsDop => Predmets.Where(p => p.Oplata.Contains("стимулир")).ToList();


            List<GroupedPredmet> GroupPredmets(List<Predmet> predmets)
            {
                var gr = predmets.GroupBy(p => new { p.Name, p.Sem });

                List<GroupedPredmet> list = new List<GroupedPredmet>();

                foreach (var p in gr)
                {
                    var g = new GroupedPredmet();
                    g.Name = p.Key.Name;
                    g.Sem = p.Key.Sem;
                    foreach (var el in p)
                    {
                        g.Groups.Add(el.Groups);
                        switch (el.Type.ToLower())
                        {
                            case "лекции":
                                g.Hours[0] = el.Vsego;
                                break;
                            case "лабораторные работы":
                                g.Hours[1] = el.Vsego;
                                break;
                            case "практические занятия":
                                g.Hours[2] = el.Vsego;
                                break;
                            case "курсовые работы/проекты":
                                g.Hours[3] = el.Vsego;
                                break;
                            case "индивидуальная работа":
                                g.Hours[4] = el.Vsego;
                                break;

                            default:
                                if (el.Name.ToLower() == "Учебная практика".ToLower())
                                    g.Hours[5] = el.Vsego;
                                else if (el.Name.ToLower() == "Производственная практика (научно-исследовательская работа)".ToLower())
                                    g.Hours[5] = el.Vsego;
                                else if (el.Name.ToLower() == "Производственная практика".ToLower())
                                    g.Hours[5] = el.Vsego;

                                else if (el.Name.ToLower() == "Дипломное проектирование (бакалавры)".ToLower())
                                    g.Hours[9] = el.Vsego;
                                else if (el.Name.ToLower() == "Преддипломная практика".ToLower())
                                    g.Hours[8] = el.Vsego;

                                else if (el.Name.ToLower() == "Защита дипломного проекта (бакалавры)".ToLower())
                                    g.Hours[9] = el.Vsego;

                                else if (el.Name.ToLower() == "Учебная практика (ознакомительная)".ToLower())
                                    g.Hours[10] = el.Vsego;
                                else if (el.Name.ToLower() == "Учебная практика (педагогическая)".ToLower())
                                    g.Hours[10] = el.Vsego;
                                else if (el.Name.ToLower() == "Производственная практика (НИР)".ToLower())
                                    g.Hours[10] = el.Vsego;
                                else if (el.Name.ToLower() == "Учебная практика (педагогическая)".ToLower())
                                    g.Hours[10] = el.Vsego;
                                else if (el.Name.ToLower() == "Производственная практика (проектно-технологическая)".ToLower())
                                    g.Hours[10] = el.Vsego;
                                else if (el.Name.ToLower() == "Производственная практика (НИР)".ToLower())
                                    g.Hours[10] = el.Vsego;
                                else if (el.Name.ToLower() == "Производственная практика (проектно-технологическая)".ToLower())
                                    g.Hours[10] = el.Vsego;
                                else if (el.Name.ToLower() == "Подготовка магистерской ВКР".ToLower())
                                    g.Hours[12] = el.Vsego;

                                else if (el.Name.ToLower() == "Производственная практика (преддипломная)".ToLower())
                                    g.Hours[10] = el.Vsego;
                                else if (el.Name.ToLower() == "Защита магистерской ВКР".ToLower())
                                    g.Hours[12] = el.Vsego;
                                break;
                        }
                    }
                    list.Add(g);
                }

                return list;
            }
            public List<GroupedPredmet> GroupedPredmets { 
                get
                {
                    return GroupPredmets(PredmetsShtat);
                }
            }
            public List<GroupedPredmet> GroupedPredmetsDop
            {
                get
                {
                    return GroupPredmets(PredmetsDop);
                }
            }


            public string Trudoustr { get; set; } = string.Empty;
            public string Dolzhnost { get; set; } = string.Empty;
            public string DolzhnostRP { get; set; } = string.Empty;
            public string NameRP { get; set; } = string.Empty;
            public string Zvanie { get; set; } = string.Empty;
            public string Stepen { get; set; } = string.Empty;
            public int Stazh { get; set; } = -1;

            public string StazhStr { get => Stazh.ToString(); set
                {
                    int val;
                    if (int.TryParse(value, out val))
                        Stazh = val;

                    OnPropertyChanged(nameof(Stazh));
                    OnPropertyChanged(nameof(StazhStr));
                }
            }

            public override string ToString()
            {
                return Name;
            }
            //----- not parsed



            public Plan()
            {
                Foses.CollectionChanged += (sender, e) =>
                {
                    UpdateFoses();
                };
                NauchOrg.CollectionChanged += (sender, e) =>
                {
                    UpdateNauchOrg();
                };
                NauchIssl.CollectionChanged += (sender, e) =>
                {
                    UpdateNauchIssl();
                };
                NauchMetod.CollectionChanged += (sender, e) =>
                {
                    UpdateNauchMetod();
                };
            }

            public class Rabota : Base
            {
                Action update;
                public Rabota(Action update)
                {
                    this.update = update;
                }
                string name;
                public string Name { get => name; set { name = value; update?.Invoke(); } }
                string hours;
                public string Hours { get => hours; set { hours = value; update?.Invoke(); } }
                string period;
                public string Period { get => period; set { period = value; update?.Invoke(); } }
            }

            #region foses

            List<(string Fos, decimal Hours, string Period)> foses = new List<(string Fos, decimal Hours, string Period)>();

            
            bool CheckPeriod(string period)
            {
                if (period.ToLower() == "осень" ||
                    period.ToLower() == "весна")
                    return true;
                return false;
            }
            public void UpdateFoses()
            {
                foses.Clear();
                foreach (var fos in Foses)
                {
                    if(fos.Period != null && CheckPeriod(fos.Period))
                        foses.Add((fos.Name, Convert.ToDecimal(fos.Hours), fos.Period));
                }
            }
            public ObservableCollection<Rabota> Foses
            {
                get;
            } = new ObservableCollection<Rabota>();
            Rabota selectedFos;
            public Rabota SelectedFos
            {
                get => selectedFos;
                set
                {
                    selectedFos = value;
                    OnPropertyChanged(nameof(SelectedFos));
                }
            }
            #endregion

            #region nauchorg

            List<(string Nauchorg, decimal Hours, string Period)> nauchorg = new List<(string Nauchorg, decimal Hours, string Period)>();

            public void UpdateNauchOrg()
            {
                nauchorg.Clear();
                foreach (var r in NauchOrg)
                {
                    if (r.Period != null && CheckPeriod(r.Period))
                        nauchorg.Add((r.Name, Convert.ToDecimal(r.Hours), r.Period));
                }
            }
            public ObservableCollection<Rabota> NauchOrg
            {
                get;
            } = new ObservableCollection<Rabota>();
            Rabota selectedNauchOrg;
            public Rabota SelectedNauchOrg
            {
                get => selectedNauchOrg;
                set
                {
                    selectedNauchOrg = value;
                    OnPropertyChanged(nameof(SelectedNauchOrg));
                }
            }
            #endregion


            #region nauchissl

            List<(string Nauchissl, decimal Hours, string Period)> nauchissl = new List<(string Nauchorg, decimal Hours, string Period)>();

            public void UpdateNauchIssl()
            {
                nauchissl.Clear();
                foreach (var r in NauchIssl)
                {
                    if (r.Period != null && CheckPeriod(r.Period))
                        nauchissl.Add((r.Name, Convert.ToDecimal(r.Hours), r.Period));
                }
            }
            public ObservableCollection<Rabota> NauchIssl
            {
                get;
            } = new ObservableCollection<Rabota>();
            Rabota selectedNauchIssl;
            public Rabota SelectedNauchIssl
            {
                get => selectedNauchIssl;
                set
                {
                    selectedNauchIssl = value;
                    OnPropertyChanged(nameof(SelectedNauchIssl));
                }
            }
            #endregion


            #region nauchissl

            List<(string Nauchmetod, decimal Hours, string Period)> nauchmetod = new List<(string Nauchorg, decimal Hours, string Period)>();

            public void UpdateNauchMetod()
            {
                nauchmetod.Clear();
                foreach (var r in NauchMetod)
                {
                    if (r.Period != null && CheckPeriod(r.Period))
                        nauchmetod.Add((r.Name, Convert.ToDecimal(r.Hours), r.Period));
                }
            }
            public ObservableCollection<Rabota> NauchMetod
            {
                get;
            } = new ObservableCollection<Rabota>();
            Rabota selectedNauchMetod;
            public Rabota SelectedNauchMetod
            {
                get => selectedNauchMetod;
                set
                {
                    selectedNauchMetod = value;
                    OnPropertyChanged(nameof(SelectedNauchMetod));
                }
            }
            #endregion
        }



        class PlansData : Base
        {
            public Dictionary<string, Plan> Plans { get; } = new Dictionary<string, Plan>();

            public IEnumerable<Plan> Prepods => Plans.Values;


            Plan selectedPlan;
            public Plan SelectedPlan
            {
                get => selectedPlan; set
                {
                    selectedPlan = value;
                    OnPropertyChanged(nameof(SelectedPlan));
                }
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

        private void button2_Click(object sender, EventArgs e)
        {
            Word.Application wordApp = new Microsoft.Office.Interop.Word.Application { Visible = true };
            Word.Document template = wordApp.Documents.Open(
                System.AppDomain.CurrentDomain.BaseDirectory + @"\Document.docx",
                ReadOnly: false, Visible: true);
            template.Activate();
            template.SaveAs2(System.AppDomain.CurrentDomain.BaseDirectory + @"\Планы\ИТИВС 2021 Индивидуальный план Глинкина Максима Олеговича.docx");







            Word.Document doc = wordApp.Documents.Open(
                System.AppDomain.CurrentDomain.BaseDirectory + @"\Планы\ИТИВС 2021 Индивидуальный план Глинкина Максима Олеговича.docx",
                ReadOnly: false, Visible: true);
            doc.Activate();
            doc.SelectAllEditableRanges();


            FindAndReplace(wordApp, "%dolzhnost%", "ассистент");
            FindAndReplace(wordApp, "%dolzhrp%", "ассистента");
            FindAndReplace(wordApp, "%fiorp%", "Глинкина Максима Олеговича");
            doc.Save();
        }
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
    }


    class Base : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
