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
                                    3,
                                    ref choice);
            }
        }

        private bool Check(string selectedPath)
        {
            DirectoryInfo dir = new DirectoryInfo(selectedPath);
            List<string> listFiles = new List<string>();
            List<string> listTextr = new List<string>() { "ChangeList", "Таблицы сигналов ТС, команд ТУ и ОТУ", "Мнемосхема перегона", "Мнемосхема станции", "Список команд управления (ТУ)", "Список команд управления и ответственных команд" };

            foreach (FileInfo fl in dir.GetFiles())
            {
                switch (fl.Extension)
                {
                    case ".xls":
                        if (fl.Name.Contains("ChangeList", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("ChangeList");
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
                        if (fl.Name.Contains("Tu", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Список команд управления (ТУ)");
                        }
                        else if (fl.Name.Contains("Otu", StringComparison.OrdinalIgnoreCase))
                        {
                            listFiles.Add("Список команд управления и ответственных команд");
                        }
                        break;
                }
            }
            listFiles = listTextr.Except(listFiles.Distinct().ToList()).ToList();
            if (listFiles.Count > 0)
                return AskWindow(listFiles);
            else
                return true;

        }
    }
}