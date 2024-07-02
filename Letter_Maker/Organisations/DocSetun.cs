using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WinForms = System.Windows.Forms;

namespace Letter_Maker.Organisations
{
    class DocSetun : Document
    {
        internal DocSetun(FolderBrowserDialog folderBrowserDialog, List<string> choice)
        {
            if (folderBrowserDialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                if (Check(folderBrowserDialog.SelectedPath))
                    MakeDocument(   folderBrowserDialog.SelectedPath,
                                    organisationList.Setun,
                                    ref choice);
            }
        }

        private bool Check(string selectedPath)
        {
            DirectoryInfo dir = new DirectoryInfo(selectedPath);
            List<string> listFiles = new List<string>();
            List<string> listSetun = new List<string>() {   "Список изменений сигналов состояния напольных устройств",
                                                            "Список сигналов состояния напольных устройств", 
                                                            "Список групп сигналов состояния напольных устройств", 
                                                            "Мнемосхема перегона",
                                                            "Мнемосхема станции",
                                                            "Таблица команд ТУ",// xlsx по названию ТУ/ОТУ
                                                            "Таблица команд ОТУ"};// или

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
                            listFiles.Add("Список сигналов состояния напольных устройств");
                        }
                        break;
                    case ".csv":
                        if (fl.Name.Contains("GroupSignalList", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Список групп сигналов состояния напольных устройств");
                        }
                        break;

                    case ".jpg":
                    case ".png":
                        if (fl.Name.Contains("Stages", StringComparison.OrdinalIgnoreCase) || fl.Name.Contains("Перегон", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Мнемосхема перегона");
                        }
                        else if (fl.Name.Contains("Station", StringComparison.OrdinalIgnoreCase) || fl.Name.Contains("Станция", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Мнемосхема станции");
                        }
                        break;
                }
            }
            listFiles = listSetun.Except(listFiles.Distinct().ToList()).ToList();
            if (listFiles.Count > 0)
                return WindowOfClarify(listFiles);
            else
                return true;

        }
    }
}
