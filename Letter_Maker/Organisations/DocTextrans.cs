using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using WinForms = System.Windows.Forms;

namespace Letter_Maker.Organisations
{
    internal class DocTextrans : Document
    {
        internal DocTextrans(WinForms.FolderBrowserDialog folderBrowserDialog, List<string> choice)
        {
            if (folderBrowserDialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                if (Check(folderBrowserDialog.SelectedPath))
                    MakeDocument(   folderBrowserDialog.SelectedPath,
                                    organisationList.Textrans,
                                    ref choice);
            }
        }

        private bool Check(string selectedPath)
        {
            DirectoryInfo dir = new DirectoryInfo(selectedPath);
            List<string> listFiles = new List<string>();
            List<string> listTextr = new List<string>() {   
                                                            //xls
                                                            "Список изменений сигналов состояния напольных устройств",
                                                            //xlsx
                                                            "Таблица ТС (ТУ и ОТУ)",
                                                            //jpg,png
                                                            "Штамп",
                                                            "Мнемосхема станции",
                                                            "Мнемосхема перегона"
                                                        };

            foreach (FileInfo fl in dir.GetFiles())
            {
                switch (fl.Extension)
                {
                    case ".xls":
                        if (fl.Name.Contains("ChangeList", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Список изменений сигналов состояния напольных устройств");
                        }
                    break;
                    case ".xlsx":
                        if (fl.Name.Contains("Таблица ", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Таблица ТС (ТУ и ОТУ)");
                        }
                        break;
                    case ".jpg":
                    case ".png":
                        if (fl.Name.Contains("Штамп", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Штамп");
                        }
                        else if (fl.Name.Contains("Station", StringComparison.OrdinalIgnoreCase) || fl.Name.Contains("Станция", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Мнемосхема станции");
                        }
                        else if (fl.Name.Contains("Stages", StringComparison.OrdinalIgnoreCase) || fl.Name.Contains("Перегон", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Мнемосхема перегона");
                        }                        
                        break;
                    default:
                        break;
                }
            }
            listFiles = listTextr.Except(listFiles.Distinct().ToList()).ToList();
            if (listFiles.Count > 0)
                return WindowOfClarify(listFiles);
            else
                return true;

        }
    }
}