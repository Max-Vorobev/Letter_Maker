using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Xml.Serialization;
using Letter_Maker.Organisations;
using WinForms = System.Windows.Forms;



namespace Letter_Maker
{
    public partial class MainWindow : System.Windows.Window
    {
        Author author = new Author();


        public MainWindow()
        {

            InitializeComponent();

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
            Author_Choise.SelectedIndex = 0;
            RailRoad_Choise.ItemsSource = listRailRoad;
            RailRoad_Choise.SelectedIndex = 0;
        }

        /// Метод, в котором вызывается диалоговое окно для выбора папки, где лежат файлы
        /// <returns> Номер нажатой кнопки </returns>
        public WinForms.FolderBrowserDialog Folder_choice()
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
                dc.MakeDocument(foulder.SelectedPath,0,ref list);
            }
        }

        private void kit_table_Click(object sender, RoutedEventArgs e)
        {
            DocKit kit = new DocKit(Folder_choice(),
                                    new List<string> {  author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                    author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
                                                    RailRoad_Choise.SelectedItem.ToString(),
                                                    Station_Name.Text});
        }



        private void setun_table_Click(object sender, RoutedEventArgs e)
        {
            DocSetun setun = new DocSetun(  Folder_choice(),
                                            new List<string> {  author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                            author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
                                                            RailRoad_Choise.SelectedItem.ToString(),
                                                            Station_Name.Text});

        }

        private void tex_tranc_table_Click(object sender, RoutedEventArgs e)
        {
            DocTextrans textrans = new DocTextrans( Folder_choice(),
                                                    new List<string> {  author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                                    author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
                                                                    RailRoad_Choise.SelectedItem.ToString(),
                                                                    Station_Name.Text});
        }
        private void adk_table_Click(object sender, RoutedEventArgs e)
        {
            DocADK adk = new DocADK(Folder_choice(),
                                    new List<string> {  author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                                    author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
                                                                    RailRoad_Choise.SelectedItem.ToString(),
                                                                    Station_Name.Text});
        }

        private void yug_rkp_table_Click(object sender, RoutedEventArgs e)
        {
            DocYugRkp yugRkp = new DocYugRkp(  Folder_choice(),
                                        new List<string> {  author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                                    author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
                                                                    RailRoad_Choise.SelectedItem.ToString(),
                                                                    Station_Name.Text});
        }

        private void yug_krug_table_Click(object sender, RoutedEventArgs e)
        {
            DocYugKrug yugKrug = new DocYugKrug(Folder_choice(),
                                        new List<string> {  author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                                    author.spis.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
                                                                    RailRoad_Choise.SelectedItem.ToString(),
                                                                    Station_Name.Text});
        }

    }
}
