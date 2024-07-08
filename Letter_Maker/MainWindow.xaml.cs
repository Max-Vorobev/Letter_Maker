using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Windows;
using System.Xml.Serialization;
using Letter_Maker.Organisations;
using static System.Net.Mime.MediaTypeNames;
using WinForms = System.Windows.Forms;



namespace Letter_Maker
{
    public partial class MainWindow : System.Windows.Window
    {
        Author author = new Author();

        public MainWindow()
        {
            InitializeComponent();
            int startPosition = 0;

            XmlSerializer formatter = new XmlSerializer(typeof(Author[]));

            using (FileStream fs = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "\\Template\\0_Authors.xml", FileMode.OpenOrCreate))
            {
                Author[]? newpeople = formatter.Deserialize(fs) as Author[];

                if (newpeople != null)
                {
                    foreach (Author person in newpeople)
                    {
                        author.spis.Add(person.authorName, person.phNumber);
                    }
                    author.Sort();
                }
                else
                {
                    MessageBox.Show("Нет списка авторов");
                }

            }

            List<string> listRailRoad = new List<string> { "Горьковской", "Забайкальской", "Московской", "Октябрьской", "Приволжской", "Северной", "Северо-Кавказской" };

            Author_Choise.ItemsSource = author.spis.Keys.Select(key => key.Split(' ').First());
            Author_Choise.SelectedIndex = startPosition;
            RailRoad_Choise.ItemsSource = listRailRoad;
            RailRoad_Choise.SelectedIndex = startPosition;
        }

        /// Метод, в котором вызывается диалоговое окно для выбора папки, где лежат файлы
        /// <returns> Номер нажатой кнопки </returns>
        private WinForms.FolderBrowserDialog Folder_choice()
        {
            WinForms.FolderBrowserDialog dialog = new WinForms.FolderBrowserDialog();
            dialog.InitialDirectory = "Z:\\Станции";
            return dialog;
        }

        private void only_table_Click(object sender, RoutedEventArgs e)
        {
            var foulder = Folder_choice();
            if (foulder.ShowDialog() == WinForms.DialogResult.OK)
            {
                Document dc = new Document();
                List<string> list = new List<string>();
                dc.MakeDocument(foulder.SelectedPath,Document.organisationList.Table,ref list);
            }
        }
        /// Метод для каждой отдельной организации
        private void kit_table_Click(object sender, RoutedEventArgs e)
        {
            if (WindowOfEmptiness(Station_Name.Text))
            {
                DocKit kit = new DocKit(Folder_choice(),
                                        new List<string> {  author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                        author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
                                                        RailRoad_Choise.SelectedItem.ToString(),
                                                        Station_Name.Text});
            }
        }



        private void setun_table_Click(object sender, RoutedEventArgs e)
        {
            if (WindowOfEmptiness(Station_Name.Text))
            {
                DocSetun setun = new DocSetun(Folder_choice(),
                                            new List<string> {  author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                                author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
                                                                RailRoad_Choise.SelectedItem.ToString(),
                                                                Station_Name.Text});
            }

        }

        private void tex_tranc_table_Click(object sender, RoutedEventArgs e)
        {
            if (WindowOfEmptiness(Station_Name.Text))
            {
                DocTextrans textrans = new DocTextrans(Folder_choice(),
                                                    new List<string> {  author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                                        author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
                                                                        RailRoad_Choise.SelectedItem.ToString(),
                                                                        Station_Name.Text});
            }
        }
        private void adk_table_Click(object sender, RoutedEventArgs e)
        {
            if (WindowOfEmptiness(Station_Name.Text))
            {
                DocADK adk = new DocADK(Folder_choice(),
                                    new List<string> {  author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                        author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
                                                        RailRoad_Choise.SelectedItem.ToString(),
                                                        Station_Name.Text});
            }
        }

        private void yug_rkp_table_Click(object sender, RoutedEventArgs e)
        {
            if (WindowOfEmptiness(Station_Name.Text))
            {
                DocYugRkp yugRkp = new DocYugRkp(Folder_choice(),
                                        new List<string> {  author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                            author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
                                                            RailRoad_Choise.SelectedItem.ToString(),
                                                            Station_Name.Text});
            }
        }

        private void yug_krug_table_Click(object sender, RoutedEventArgs e)
        {
            if (WindowOfEmptiness(Station_Name.Text))
            {
                DocYugKrug yugKrug = new DocYugKrug(Folder_choice(),
                                        new List<string> {  author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                            author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
                                                            RailRoad_Choise.SelectedItem.ToString(),
                                                            Station_Name.Text});
            }
        }

        /// Метод вызова окна для уточнения - где имя станции?
        public bool WindowOfEmptiness(string stName)
        {
            if (string.IsNullOrEmpty(stName))
            {
                var result = System.Windows.MessageBox.Show(
                                                    "Не хватет имени станции\n\nПродолжить?",
                                                    "ВНИМАНИЕ!!!",
                                                    MessageBoxButton.YesNo,
                                                    MessageBoxImage.Warning);
                return result == MessageBoxResult.Yes;
            }
            return true;
        }
    }
}
