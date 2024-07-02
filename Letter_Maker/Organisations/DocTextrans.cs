using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using WinForms = System.Windows.Forms;

namespace Letter_Maker.Organisations
{
    class DocTextrans : Document
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
            List<string> listTextr = new List<string>() {   "Список изменений сигналов состояния напольных устройств",
                                                            "Мнемосхема перегона",
                                                            "Мнемосхема станции",
                                                            "Таблицы сигналов ТС, команд ТУ и ОТУ", //signal list
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
                    case ".xlsx":
                        if (fl.Name.Contains("signallist", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Таблицы сигналов ТС, команд ТУ и ОТУ");
                        }
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