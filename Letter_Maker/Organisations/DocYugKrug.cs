﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WinForms = System.Windows.Forms;

namespace Letter_Maker.Organisations
{
    internal class DocYugKrug : Document
    {
        internal DocYugKrug(WinForms.FolderBrowserDialog folderBrowserDialog, List<string> choice)
        {
            if (folderBrowserDialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                if (Check(folderBrowserDialog.SelectedPath))
                    MakeDocument(   folderBrowserDialog.SelectedPath,
                                    organisationList.YugKrug,
                                    ref choice);
            }
        }

        private bool Check(string selectedPath)
        {
            DirectoryInfo dir = new DirectoryInfo(selectedPath);
            List<string> listFiles = new List<string>();
            List<string> listYugKrug = new List<string>() {   "Список изменений сигналов состояния напольных устройств",
                                                            "Список сигналов состояния напольных устройств",
                                                            "Список групп сигналов состояния напольных устройств",
                                                            "Мнемосхема перегона",
                                                            "Мнемосхема станции",
                                                            "Таблица команд ТУ",// xlsx по названию ТУ/ОТУ
                                                            "Таблица команд ТУ/ОТУ"};// или

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
                    case ".xlsx":
                        if (fl.Name.Contains("Таблица команд ТУ", StringComparison.OrdinalIgnoreCase))
                        {
                            if (fl.Name.Contains("ОТУ", StringComparison.OrdinalIgnoreCase))
                            {
                                listFiles.Add("Таблица команд ТУ/ОТУ");
                            }
                            listFiles.Add("Таблица команд ТУ");
                        }
                        break;
                }
            }
            listFiles = listYugKrug.Except(listFiles.Distinct().ToList()).ToList();
            if (listFiles.Count > 0)
                return WindowOfClarify(listFiles);
            else
                return true;

        }
    }
}