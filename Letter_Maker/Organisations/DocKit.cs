using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinForms = System.Windows.Forms;

namespace Letter_Maker.Organisations
{
    internal class DocKit : Document
    {
        List<string> listKit = new List<string>() {     "Список изменений сигналов состояния напольных устройств",
                                                        "Список сигналов состояния напольных объектов",
                                                        "Список сигналов состояния напольных устройств",
                                                        "Список сигналов состояния устройств УВК РА",
                                                        "Список сигналов состояния устройств УСО",
                                                        "Список сигналов состояния устройств КСУ",
                                                        "Мнемосхема перегона",
                                                        "Мнемосхема станции",
                                                        "Мнемосхема каналов УСО",
                                                        "Мнемосхема шкафа УВК" };
        public DocKit(WinForms.FolderBrowserDialog folderBrowserDialog, List<string> choice)
        {
            if (folderBrowserDialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                if (Check(folderBrowserDialog.SelectedPath))
                    MakeDocument(folderBrowserDialog.SelectedPath,
                                    organisationList.Kit,
                                    ref choice);
            }
        }
        public bool Check(string selectedPath)
        {
            DirectoryInfo dir = new DirectoryInfo(selectedPath);
            List<string> listFiles = new List<string>();
            


            foreach (FileInfo fl in dir.GetFiles())
            {
                switch (fl.Extension)
                {
                    case ".xls":
                        if (fl.Name.Contains("ChangeList", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Список изменений сигналов состояния напольных устройств");
                        }
                        else if (fl.Name.Contains("SignalList", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Список сигналов состояния напольных объектов");
                        }
                        break;
                    case ".xml":
                        if (fl.Name.Contains("_full", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Список сигналов состояния напольных устройств");
                        }
                        else if (fl.Name.Contains("uvk-data_dk", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Список сигналов состояния устройств УВК РА");
                        }
                        else if (fl.Name.Contains("usoCh-data_dk", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Список сигналов состояния устройств УСО");
                        }
                        else if (fl.Name.Contains("ksuDiag", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Список сигналов состояния устройств КСУ");
                        }
                        break;
                    case ".jpg":
                    case ".png":
                        if (fl.Name.Contains("Stages", StringComparison.OrdinalIgnoreCase) || fl.Name.Contains("Перегон", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Мнемосхема перегона");
                        }
                        else if (fl.Name.Contains("Station", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Мнемосхема станции");
                        }
                        else if (fl.Name.Contains("Uso", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Мнемосхема каналов УСО");
                        }
                        else if (fl.Name.Contains("Uvk", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Мнемосхема шкафа УВК");
                        }
                        break;
                }
            }
            listFiles = listKit.Except(listFiles.Distinct().ToList()).ToList();
            if (listFiles.Count > 0)
                return WindowOfClarify(listFiles);
            else
                return true;

        }
    }
}
