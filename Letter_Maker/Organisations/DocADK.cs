using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WinForms = System.Windows.Forms;

namespace Letter_Maker.Organisations
{
    class DocADK : Document
    {
        internal DocADK(WinForms.FolderBrowserDialog folderBrowserDialog, List<string> choice)
        {
            if (folderBrowserDialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                if (Check(folderBrowserDialog.SelectedPath))
                    MakeDocument(   folderBrowserDialog.SelectedPath,
                                    4,
                                    ref choice);
            }
        }

        private bool Check(string selectedPath)
        {
            DirectoryInfo dir = new DirectoryInfo(selectedPath);
            List<string> listFiles = new List<string>();
            List<string> listADK = new List<string>() {   "Мнемосхема каналов УСО",
                                                            "Мнемосхема шкафа УВК",
                                                            "ChangeList",
                                                            "SignalList",
                                                            "Мнемосхема перегона",
                                                            "Мнемосхема станции",
                                                            "Список сигналов состояния напольных устройств",
                                                            "Список сигналов состояния устройств КСУ",
                                                            "Список сигналов состояния устройств УСО",
                                                            "Список сигналов состояния устройств УВК РА"};

            foreach (FileInfo fl in dir.GetFiles())
            {
                switch (fl.Extension)
                {
                    case ".xls":
                        if (fl.Name.Contains("ChangeList", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("ChangeList");
                        }
                        else if (fl.Name.Contains("SignalList", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("SignalList");
                        }
                        break;
                    case ".xml":
                        if (fl.Name.Contains("full", StringComparison.OrdinalIgnoreCase))
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
            listFiles = listADK.Except(listFiles.Distinct().ToList()).ToList();
            if (listFiles.Count > 0)
                return AskWindow(listFiles);
            else
                return true;

        }
    }
}
