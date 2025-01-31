using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WinForms = System.Windows.Forms;

namespace Letter_Maker.Organisations
{
    internal class DocASDK : Document
    {
        internal DocASDK(WinForms.FolderBrowserDialog folderBrowserDialog, List<string> choice)
        {
            if (folderBrowserDialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                if (Check(folderBrowserDialog.SelectedPath))
                    MakeDocument(folderBrowserDialog.SelectedPath,
                                    organisationList.ASDK,
                                    ref choice);
            }
        }

        private bool Check(string selectedPath)
        {
            DirectoryInfo dir = new DirectoryInfo(selectedPath);
            List<string> listFiles = new List<string>();
            List<string> listASDK = new List<string>() {     
                                                            //xls
                                                            "Список изменений сигналов состояния напольных устройств",
                                                            "Список сигналов состояния напольных объектов",
                                                            //xml
                                                            "Список сигналов состояния напольных устройств",
                                                            "Список сигналов состояния устройств УВК РА",
                                                            "Список сигналов состояния устройств УСО",
                                                            "Список сигналов состояния устройств КСУ",
                                                            //jpg,png
                                                            "Штамп",
                                                            "Мнемосхема станции",
                                                            "Мнемосхема перегона",
                                                            "Мнемосхема шкафа УВК",
                                                            "Мнемосхема каналов УСО"
                                                      };

            foreach (FileInfo fl in dir.GetFiles())
            {
                switch (fl.Extension.ToLower())
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
                        if (fl.Name.Contains("Штамп", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Штамп");
                        }
                        else if (fl.Name.Contains("Station", StringComparison.OrdinalIgnoreCase) || fl.Name.Contains("Станци", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Мнемосхема станции");
                        }
                        else if (fl.Name.Contains("Stages", StringComparison.OrdinalIgnoreCase) || fl.Name.Contains("Перегон", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Мнемосхема перегона");
                        }
                        else if (fl.Name.Contains("Uvk", StringComparison.OrdinalIgnoreCase) || fl.Name.Contains("УВК", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Мнемосхема шкафа УВК");
                        }
                        else if (fl.Name.Contains("Uso", StringComparison.OrdinalIgnoreCase) || fl.Name.Contains("УСО", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Мнемосхема каналов УСО");
                        }
                        break;
                    default:
                        break;
                }
            }
            listFiles = listASDK.Except(listFiles.Distinct().ToList()).ToList();
            if (listFiles.Count > 0)
                return WindowOfClarify(listFiles);
            else
                return true;

        }
    }
}
