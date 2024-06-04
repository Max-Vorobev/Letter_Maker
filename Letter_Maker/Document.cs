using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using Word = Microsoft.Office.Interop.Word;

namespace Letter_Maker
{
    internal class Document : MainWindow
    {

        public Document() { }

        /// Метод для создания документа с таблицей описывающей передаваемые файлы

        public void MakeDocument(string theWay, int option,ref List<string> Aut_Ch)
        {
            Word.Application fileOpen = new Word.Application();
            Word.Document? dc = null;
            string? fName = null;
            int paragraphPos = 1;
            switch (option)
            {
                case 0:// Таблица
                    object missing = System.Reflection.Missing.Value;
                    dc = fileOpen.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                    fName = "\\таблица.doc";
                    fileOpen.Visible = false;
                    dc.Activate();
                    dc.PageSetup.LeftMargin = (float)50;
                    dc.PageSetup.TopMargin = (float)50;
                    break;
                case 1:// КИТ
                    dc = fileOpen.Documents.Open(AppDomain.CurrentDomain.BaseDirectory + "\\Template\\Kit.doc", ReadOnly: false);
                    fName = $"\\{DateTime.Now.ToString("yyyy.MM.dd")} М.А.Еремин С.Э.Усачеву - Материалы для адаптации ст. " + Aut_Ch[3] + $".doc";
                    paragraphPos = 19;
                    fileOpen.Visible = false;
                    dc.Activate();
                    listOfChage(ref fileOpen,ref Aut_Ch);
                    break;
                case 2:// Сетунь
                    dc = fileOpen.Documents.Open(AppDomain.CurrentDomain.BaseDirectory + "\\Template\\Setun.doc", ReadOnly: false);
                    fName = $"\\{DateTime.Now.ToString("yyyy.MM.dd")} М.А.Еремин П.В.Бармину - Материалы для адаптации ст. " + Aut_Ch[3] +".doc";
                    paragraphPos = 19;
                    fileOpen.Visible = false;
                    dc.Activate();
                    listOfChage(ref fileOpen, ref Aut_Ch);
                    break;
                case 3: // Техтранс
                    dc = fileOpen.Documents.Open(AppDomain.CurrentDomain.BaseDirectory + "\\Template\\Textrans.doc", ReadOnly: false);
                    fName = $"\\{DateTime.Now.ToString("yyyy.MM.dd")} М.А.Ерёмин А.С.Павлову - Материалы для адаптации ст. "+ Aut_Ch[3] + ".doc";
                    paragraphPos = 23;
                    fileOpen.Visible = false;
                    dc.Activate();
                    listOfChage(ref fileOpen, ref Aut_Ch);
                    break;
                case 4: // АДК СЦБ
                    dc = fileOpen.Documents.Open(AppDomain.CurrentDomain.BaseDirectory + "\\Template\\ADK.doc", ReadOnly: false);
                    fName = $"\\{DateTime.Now.ToString("yyyy.MM.dd")} М.А.Еремин С.А.Панову - Материалы для адаптации ПО ст. " + Aut_Ch[3] + ".doc";
                    paragraphPos = 18;
                    fileOpen.Visible = false;
                    dc.Activate();
                    listOfChage(ref fileOpen, ref Aut_Ch);
                    break;
                /////////////////////////// !!!!!!!!!!!!!!!!!! Поправить Параграфы и Шаблоны !!!!!!!!!!!!!!!!!!
                case 5: // Юг РКП
                    dc = fileOpen.Documents.Open(AppDomain.CurrentDomain.BaseDirectory + "\\Template\\YugKrug.doc", ReadOnly: false);
                    fName = $"\\{DateTime.Now.ToString("yyyy.MM.dd")}  М.А.Еремин Л.П.Кузнецову - Материалы для адаптации ПО ст. " + Aut_Ch[3] + ".doc";
                    paragraphPos = 18;
                    fileOpen.Visible = false;
                    dc.Activate();
                    listOfChage(ref fileOpen, ref Aut_Ch);
                    break;
                case 6: // Юг, тот что пел про Владимирский Централ
                    dc = fileOpen.Documents.Open(AppDomain.CurrentDomain.BaseDirectory + "\\Template\\YugRkp.doc", ReadOnly: false);
                    fName = $"\\{DateTime.Now.ToString("yyyy.MM.dd")} М.А. Еремин В.В. Аракельяну - Материалы для адаптации ПО ст. " + Aut_Ch[3] + ".doc";
                    paragraphPos = 18;
                    fileOpen.Visible = false;
                    dc.Activate();
                    listOfChage(ref fileOpen, ref Aut_Ch);
                    break;
            }

            
            DirectoryInfo dir = new DirectoryInfo(theWay);
            int kolich = dir.GetFiles().Length;

            Word.Range wordrange = dc.Paragraphs[paragraphPos].Range;
            Object defaultTableBehavior = WdDefaultTableBehavior.wdWord9TableBehavior;
            Object autoFitBehavior = WdAutoFitBehavior.wdAutoFitWindow;

            //Добавляем таблицу и получаем объект wordtable 
            Word.Table wordtable = dc.Tables.Add(wordrange, kolich + 1, 5, ref defaultTableBehavior, ref autoFitBehavior);
            Word.Table tbl = dc.Tables[1];
            tbl.Range.Font.Size = 9;
            tbl.Range.Paragraphs.LineSpacing = 12;
            
            /// Вызов метода для формирования заголовка таблицы
            Zagalovok(tbl);

            int ch = 2; // Потому что 1 - это заголовок
            foreach (FileInfo fl in dir.GetFiles())
            {
                tbl.Rows[ch].Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                tbl.Cell(ch, 1).Range.Text = (ch - 1).ToString();
                tbl.Cell(ch, 2).Range.Text = Name(fl.Name);
                tbl.Cell(ch, 3).Range.Text = fl.Name;
                tbl.Cell(ch, 4).Range.Text = fl.Length.ToString("N0");
                tbl.Cell(ch, 5).Range.Text = fl.CreationTime.ToString("g");
                ch++;
            }
            /// сохраняем файл и закрываем его
            dc.SaveAs2(theWay + fName);
            dc.Close();
            /// делаем красиво по памяти, закрываем Word и выводим сообщение, что всё готово
            dc = null;
            fileOpen.Quit();
            fileOpen = null;
            MessageBox.Show("Файл сформирован");
        }

        /// Метод для замены текста в документе

        public void FindAndReplace(Word.Application fileOpen, object findText, object replaceWithText)
        {
            // Задаем параметры замены
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object replace = 2;
            object wrap = 1;
            
            // Добавляем форматирование подчеркивания к замененному тексту
            fileOpen.Selection.Font.Underline = WdUnderline.wdUnderlineSingle;

            // Заменяем и добавляем форматирование подчеркивания
            fileOpen.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace);

            
        }

        /// Метод для формирования заголовка таблицы
        public void Zagalovok(Word.Table table)
        {
            table.Rows[1].Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Range.Paragraphs.SpaceBefore = 3;
            table.Range.Paragraphs.SpaceAfter = 3;
            table.Rows[1].Range.Font.Bold = 1;
            /// Задаём ширину каждого столба
            table.Columns[1].Width = 15;
            table.Columns[2].Width = 130;
            table.Columns[3].Width = 230;
            table.Columns[4].Width = 60;
            table.Columns[5].Width = 80;
            /// Заголовочная часть
            table.Cell(1, 1).Range.Text = "№";
            table.Cell(1, 2).Range.Text = "Наименование документа";
            table.Cell(1, 3).Range.Text = "Наименование файла документа";
            table.Cell(1, 4).Range.Text = "Размер файла, б";
            table.Cell(1, 5).Range.Text = "Время изм. файла";
        }

        /// Метод для пакетной замены тегов на текст в документе
        void listOfChage(ref Word.Application fl, ref List<string> choise)
        {
            FindAndReplace(fl, "<date>", DateTime.Now.ToString("d"));
            FindAndReplace(fl, "<author>", choise[0]);
            FindAndReplace(fl, "<phone>", choise[1]);
            FindAndReplace(fl, "<rr>", choise[2]);
            if (choise[2] == "Октябрьской")
            {
                FindAndReplace(fl, "<okt>", "КОПИЯ:\vСлужба Ш Октябрьской\vдирекции инфраструктуры, \vНачальнику отдела развития и перспективных технологий \vП. А. Капусте\v");
                FindAndReplace(fl, "<okt_mail>", "pele1968@mail.ru");
                FindAndReplace(fl, "<cm>", ",");
                FindAndReplace(fl, "<dt>", ".");

            }
            else
            {
                FindAndReplace(fl, "<okt>", "");
                FindAndReplace(fl, "<okt_mail>", "");
                FindAndReplace(fl, "<cm>", ".");
                FindAndReplace(fl, "<dt>", "");
            }
            FindAndReplace(fl, "<station>", choise[3]);
        }

        public bool AskWindow(List<string> listFiles)
        {
            var result = System.Windows.MessageBox.Show(
                                                "Не хватетследующих файлов:\n\n" + String.Join("\n", listFiles) + "\n\nПродолжить?",
                                                "ВНИМАНИЕ!!!",
                                                MessageBoxButton.YesNo,
                                                MessageBoxImage.Warning);
            return result == MessageBoxResult.Yes;
        }

        /// Метод для формирования Наименование документа в таблице (2 столбец)
        private string Name(string fName)
        {
            if (fName.Contains(".xls"))
                if (fName.Contains("ChangeList", StringComparison.OrdinalIgnoreCase))
                {
                    return "Список изменений сигналов состояния напольных устройств";
                }
                else if (fName.Contains("SignalList", StringComparison.OrdinalIgnoreCase))
                {
                    return "Список сигналов состояния напольных устройств";
                }
                else
                    return "-";
            else if (fName.Contains(".xml"))
                if (fName.Contains("full", StringComparison.OrdinalIgnoreCase))
                {
                    return "Список сигналов состояния напольных устройств";
                }
                else if (fName.Contains("uvk-data_dk", StringComparison.OrdinalIgnoreCase))
                {
                    return "Список сигналов состояния устройств УВК РА";
                }
                else if (fName.Contains("usoCh-data_dk", StringComparison.OrdinalIgnoreCase))
                {
                    return "Список сигналов состояния устройств УСО";
                }
                else if (fName.Contains("ksuDiag", StringComparison.OrdinalIgnoreCase))
                {
                    return "ksuDiag";
                }
                else
                    return "-";
            else if (fName.Contains(".jpg") || fName.Contains(".png"))
                if (fName.Contains("Stages", StringComparison.OrdinalIgnoreCase) || fName.Contains("Перегон", StringComparison.OrdinalIgnoreCase))
                {
                    return "Мнемосхема перегона";
                }
                else if (fName.Contains("Station", StringComparison.OrdinalIgnoreCase))
                {
                    return "Мнемосхема станции";
                }
                else if (fName.Contains("Uso", StringComparison.OrdinalIgnoreCase))
                {
                    return "Мнемосхема каналов УСО";
                }
                else if (fName.Contains("Uvk", StringComparison.OrdinalIgnoreCase))
                {
                    return "Мнемосхема шкафа УВК";
                }
                else
                    return "-";
            else
                return "-";
        }
    }
}
