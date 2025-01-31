using System;
using System.Collections.Generic;
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
        rrList rrSpis = new rrList();

        public MainWindow()
        {
            InitializeComponent();
            int startPosition = 0;
            XmlSerializer formatter = new XmlSerializer(typeof(ListModel));

            try
            {
                using (FileStream fs = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "\\Template\\0_Authors.xml", FileMode.Open))
                {
                    ListModel listModel = (ListModel)formatter.Deserialize(fs);

                    if (listModel != null)
                    {
                        foreach (Author person in listModel.Authors)
                        {
                            author.spisAuthor.Add(person.Name, person.PhoneNumber);
                        }
                        

                        foreach (string org in listModel.rrLst)
                        {
                            rrSpis.spisRR.Add(org);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Нет списка авторов");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка десериализации: {ex.Message}");
            }
            rrSpis.spisRR.Sort();
            
            List<string> systemChoise = new List<string> { "Отсутствует","АБТЦ-МШ", "УРЦК"};

            Author_Choise.ItemsSource = author.spisAuthor.Keys.Select(key => key.Split(' ').First());
            Author_Choise.SelectedIndex = startPosition;
            RailRoad_Choise.ItemsSource = rrSpis.spisRR;
            RailRoad_Choise.SelectedIndex = startPosition;
            System_Choise.ItemsSource = systemChoise;
            System_Choise.SelectedIndex = startPosition;
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
                                        new List<string> {  author.spisAuthor.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                            author.spisAuthor.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
                                                            RailRoad_Choise.SelectedItem.ToString(),
                                                            Station_Name.Text,
                                                            System_Choise.SelectedItem.ToString()});
            }
        }



        private void setun_table_Click(object sender, RoutedEventArgs e)
        {
            if (WindowOfEmptiness(Station_Name.Text))
            {
                DocSetun setun = new DocSetun(Folder_choice(),
                                            new List<string> {  author.spisAuthor.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                                author.spisAuthor.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
                                                                RailRoad_Choise.SelectedItem.ToString(),
                                                                Station_Name.Text});
            }

        }

        private void tex_tranc_table_Click(object sender, RoutedEventArgs e)
        {
            if (WindowOfEmptiness(Station_Name.Text))
            {
                DocTextrans textrans = new DocTextrans(Folder_choice(),
                                                    new List<string> {  author.spisAuthor.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                                        author.spisAuthor.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
                                                                        RailRoad_Choise.SelectedItem.ToString(),
                                                                        Station_Name.Text});
            }
        }
        private void adk_scb_table_Click(object sender, RoutedEventArgs e)
        {
            if (WindowOfEmptiness(Station_Name.Text))
            {
                DocADKSCB adk = new DocADKSCB(Folder_choice(),
                                        new List<string> {  author.spisAuthor.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                            author.spisAuthor.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
                                                            RailRoad_Choise.SelectedItem.ToString(),
                                                            Station_Name.Text});
            }
        }

        private void asdk_table_Click (object sender, RoutedEventArgs e)
        {
            if (WindowOfEmptiness(Station_Name.Text))
            {
                DocASDK adk = new DocASDK(Folder_choice(),
                                        new List<string> {  author.spisAuthor.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                            author.spisAuthor.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
                                                            RailRoad_Choise.SelectedItem.ToString(),
                                                            Station_Name.Text});
            }
        }

        private void yug_rkp_table_Click(object sender, RoutedEventArgs e)
        {
            if (WindowOfEmptiness(Station_Name.Text))
            {
                DocYugRkp yugRkp = new DocYugRkp(Folder_choice(),
                                        new List<string> {  author.spisAuthor.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                            author.spisAuthor.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
                                                            RailRoad_Choise.SelectedItem.ToString(),
                                                            Station_Name.Text});
            }
        }

        private void yug_krug_table_Click(object sender, RoutedEventArgs e)
        {
            if (WindowOfEmptiness(Station_Name.Text))
            {
                DocYugKrug yugKrug = new DocYugKrug(Folder_choice(),
                                        new List<string> {  author.spisAuthor.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Key,
                                                            author.spisAuthor.FirstOrDefault(x => x.Key.StartsWith(Author_Choise.SelectedItem.ToString())).Value,
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
