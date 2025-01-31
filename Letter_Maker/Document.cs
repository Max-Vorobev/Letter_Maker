using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Linq;

namespace Letter_Maker
{
    internal class Document : MainWindow
    {

        public Document() { }

        // Перечисления организаций
        public enum organisationList
        {
            Table,
            Kit,
            Setun,
            Textrans,
            ADKSCB,
            YugRkp,
            YugKrug,
            ASDK
        }

        /// private методы 

        private static int FindParagraphWithTag(Word.Document document, string tag)
        {
            int nLevel = 0;
            foreach (Paragraph paragraph in document.Paragraphs)
            {
                Word.Range parRng = paragraph.Range;
                string sText = parRng.Text.TrimEnd('\r');
                nLevel++;
                if (sText == tag)
                    return nLevel;
            }
            

            return 1;
        }
        // Метод для замены текста в документе
        private void FindAndReplace(Word.Application fileOpen, object findText, object replaceWithText)
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

            
            // Заменяем и добавляем форматирование подчеркивания
            fileOpen.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace);


        }

        // Метод для формирования заголовка таблицы
        private void Zagalovok(Word.Table table)
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

        // Метод для пакетной замены тегов на текст в документе
        private void listOfChage(ref Word.Application fl, ref List<string> choise)
        {
            FindAndReplace(fl, "<date>", DateTime.Now.ToString("d"));
            FindAndReplace(fl, "<author>", choise[0]);
            FindAndReplace(fl, "<phone>", choise[1]);
            FindAndReplace(fl, "<rr>", choise[2]);
            FindAndReplace(fl, "<okt>", "");
            FindAndReplace(fl, "<okt_mail>", ".");
            FindAndReplace(fl, "<station>", choise[3]);
        }

        // Метод для формирования Наименование документа в таблице (2 столбец)
        private string GiveFileDiscription(string fName)
        {
            if (fName.Contains(".csv"))
                if (fName.Contains("GroupSignalList", StringComparison.OrdinalIgnoreCase))
                {
                    return "Список групп сигналов состояния напольных устройств";
                }
                else
                    return "-";
            else if (fName.Contains(".xls")) // тут же и xlsx
                if (fName.Contains("ChangeList", StringComparison.OrdinalIgnoreCase))
                {
                    return "Список изменений сигналов состояния напольных устройств";
                }
                else if (fName.Contains("SignalList", StringComparison.OrdinalIgnoreCase))
                {
                    return "Список сигналов состояния напольных устройств";
                }
                else if (fName.Contains("Таблица", StringComparison.OrdinalIgnoreCase))
                {
                    if (fName.Contains("ОТУ", StringComparison.OrdinalIgnoreCase))
                        return "Таблица ТС,ТУ и ОТУ";
                    else if (fName.Contains("ТУ", StringComparison.OrdinalIgnoreCase))
                        return "Таблица ТС и ТУ";
                    else
                        return "Таблица ТС";
                }
                else if (fName.Contains("ТУ", StringComparison.OrdinalIgnoreCase))
                {
                    if (fName.Contains("ОТУ", StringComparison.OrdinalIgnoreCase))
                    {
                        return ("Таблица команд ТУ / ОТУ");
                    }
                    else
                        return "Таблица команд ТУ";
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
                    return "Список сигналов состояния устройств КСУ";
                }
                // для АБТЦ-МШ
                else if (fName.Contains("abtcmshDiag-data_dk", StringComparison.OrdinalIgnoreCase))
                {
                    return "Список сигналов состояния устройств АБТЦ-МШ";
                }
                //для УРЦК
                else if (fName.Contains("urckUvkBrief", StringComparison.OrdinalIgnoreCase))
                {
                    return "Диагностика связей УРЦК";
                }
                else if (fName.Contains("urckProcessed-data", StringComparison.OrdinalIgnoreCase))
                {
                    return "Диагностическая информация УРЦК";
                }
                else
                    return "-";
            else if (fName.Contains(".jpg", StringComparison.OrdinalIgnoreCase) || fName.Contains(".png", StringComparison.OrdinalIgnoreCase))
                if (fName.Contains("Штамп", StringComparison.OrdinalIgnoreCase))
                {
                    return "Штамп схематического плана";
                }
                else if ((fName.Contains("Station", StringComparison.OrdinalIgnoreCase)) || (fName.Contains("Станци", StringComparison.OrdinalIgnoreCase)))
                {
                    return "Мнемосхема станции";
                }
                else if (fName.Contains("Stages", StringComparison.OrdinalIgnoreCase) || fName.Contains("Перегон", StringComparison.OrdinalIgnoreCase))
                {
                    return "Мнемосхема перегона";
                }
                else if (fName.Contains("Uvk", StringComparison.OrdinalIgnoreCase) || fName.Contains("УВК", StringComparison.OrdinalIgnoreCase))
                {
                    return "Мнемосхема шкафа УВК";
                }
                else if (fName.Contains("Uso", StringComparison.OrdinalIgnoreCase) || fName.Contains("УСО", StringComparison.OrdinalIgnoreCase))
                {
                    return "Мнемосхема каналов УСО";
                }
                else if (fName.Contains("URCK", StringComparison.OrdinalIgnoreCase) || fName.Contains("УРЦК", StringComparison.OrdinalIgnoreCase))
                {
                    Regex regex = new Regex(@"\d+-?\d*");
                    return "Таблица с данными ДИ модулей УРЦК " + regex.Match(fName);
                }
                else
                    return "-";
            else
                return "-";
        }

        /// Метод для формирования имени итогового документа
        private string MakeFileName(string aliceBob, string stationChoice)
        {
            return $"\\{DateTime.Now.ToString("yyyy.MM.dd")} {aliceBob} - Материалы для адаптации ПО ст." + stationChoice + ".doc";
        }


        /// public методы 

        /// <summary>
        /// Метод для создания документа с таблицей описывающей передаваемые файлы
        /// </summary>
        /// <param name="theWay">Путь выбранный оператором</param>
        /// <param name="option">Организация для которой будет формироваться письмо</param>
        /// <param name="Aut_Ch">Входные данные, введенные оператором</param> 
        public void MakeDocument(string theWay, organisationList option,ref List<string> Aut_Ch)
        {
            // Избегаем magic-number
            int startPosition = 1;       
            
            Word.Application fileOpen = new Word.Application();
            Word.Document? wordDocument = null;
            string? fName = null;
            int paragraphPos = startPosition;
            switch (option)
            {
                case organisationList.Kit: // АПК ДК КИТ
                    wordDocument = fileOpen.Documents.Open(AppDomain.CurrentDomain.BaseDirectory + "\\Template\\Kit.doc", ReadOnly: false);
                    fName = MakeFileName("М.А.Еремин С.Э.Усачеву", Aut_Ch[3]);
                    fileOpen.Visible = false;
                    wordDocument.Activate();
                    listOfChage(ref fileOpen,ref Aut_Ch);
                    break;
                case organisationList.ADKSCB: // АДК СЦБ ЮгПа
                    wordDocument = fileOpen.Documents.Open(AppDomain.CurrentDomain.BaseDirectory + "\\Template\\ADKSCB.doc", ReadOnly: false);
                    fName = MakeFileName("М.А.Еремин С.А.Панову", Aut_Ch[3]);
                    fileOpen.Visible = false;
                    wordDocument.Activate();
                    listOfChage(ref fileOpen, ref Aut_Ch);
                    break;
                case organisationList.ASDK: // АСДК
                    wordDocument = fileOpen.Documents.Open(AppDomain.CurrentDomain.BaseDirectory + "\\Template\\ASDK.doc", ReadOnly: false);
                    fName = MakeFileName("М.А.Еремин С.А.Аверкиеву", Aut_Ch[3]);
                    fileOpen.Visible = false;
                    wordDocument.Activate();
                    listOfChage(ref fileOpen, ref Aut_Ch);
                    break;
                case organisationList.Setun:// Сетунь
                    wordDocument = fileOpen.Documents.Open(AppDomain.CurrentDomain.BaseDirectory + "\\Template\\Setun.doc", ReadOnly: false);
                    fName = MakeFileName("М.А.Еремин П.В.Бармину", Aut_Ch[3]);
                    fileOpen.Visible = false;
                    wordDocument.Activate();
                    if (Aut_Ch[2] == "Приволжской")
                    {
                        FindAndReplace(fileOpen, "<okt>", "КОПИЯ:\vГлавному инженеру службы\vавтоматики и телемеханики\vПриволжской дирекции\vинфраструктуры\vЛ. В. Шулятьеву\v");
                        FindAndReplace(fileOpen, "<okt_mail>", ", Sh_shulyatevLV@pvrr.rzd.ru, shch6_kaurovam@pvrr.rzd.ru.");
                    }
                    else if (Aut_Ch[2] == "Октябрьской")
                    {
                        FindAndReplace(fileOpen, "<okt>", "КОПИЯ:\vСлужба Ш Октябрьской\vдирекции инфраструктуры, \vНачальнику отдела развития и перспективных технологий \vП. А. Капусте\v");
                        FindAndReplace(fileOpen, "<okt_mail>", ", sh_kapusta@orw.rzd.ru, pele1968@mail.ru.");
                    }
                    listOfChage(ref fileOpen, ref Aut_Ch);
                    break;
                case organisationList.YugRkp: // ДЦ Юг с Ркп
                    wordDocument = fileOpen.Documents.Open(AppDomain.CurrentDomain.BaseDirectory + "\\Template\\YugRkp.doc", ReadOnly: false);
                    fName = MakeFileName("М.А.Еремин Л.П.Кузнецову", Aut_Ch[3]);
                    fileOpen.Visible = false;
                    wordDocument.Activate();
                    listOfChage(ref fileOpen, ref Aut_Ch);
                    break;
                case organisationList.YugKrug: // ДЦ ЮГ Круг
                    wordDocument = fileOpen.Documents.Open(AppDomain.CurrentDomain.BaseDirectory + "\\Template\\YugKrug.doc", ReadOnly: false);
                    fName = MakeFileName("М.А. Еремин В.В. Аракельяну", Aut_Ch[3]);
                    fileOpen.Visible = false;
                    wordDocument.Activate();
                    listOfChage(ref fileOpen, ref Aut_Ch);
                    break;
                case organisationList.Textrans: // ДЦ Тракт - "Техтрнас"
                    wordDocument = fileOpen.Documents.Open(AppDomain.CurrentDomain.BaseDirectory + "\\Template\\Textrans.doc", ReadOnly: false);
                    fName = MakeFileName("М.А.Ерёмин А.С.Павлову", Aut_Ch[3]);
                    fileOpen.Visible = false;
                    wordDocument.Activate();
                    if (Aut_Ch[2] == "Октябрьской")
                    {
                        FindAndReplace(fileOpen, "<okt>", "КОПИЯ:\vСлужба Ш Октябрьской\vдирекции инфраструктуры, \vНачальнику отдела развития и перспективных технологий \vП. А. Капусте\v");
                        FindAndReplace(fileOpen, "<okt_mail>", ", sh_kapusta@orw.rzd.ru, pele1968@mail.ru.");
                    }
                    listOfChage(ref fileOpen, ref Aut_Ch);
                    break;
                case organisationList.Table: // просто таблица
                    object missing = System.Reflection.Missing.Value;
                    wordDocument = fileOpen.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                    fName = "\\таблица.doc";
                    fileOpen.Visible = false;
                    wordDocument.Activate();
                    wordDocument.PageSetup.LeftMargin = (float)50;
                    wordDocument.PageSetup.TopMargin = (float)50;
                    break;
                default:
                    var result = System.Windows.MessageBox.Show("Случилось страшное");
                    break;
            }

            
            DirectoryInfo dir = new DirectoryInfo(theWay);
            int kolich = dir.GetFiles().Length;
            Word.Range wordrange = wordDocument.Paragraphs[FindParagraphWithTag(wordDocument,"<table>")].Range;
            Object defaultTableBehavior = WdDefaultTableBehavior.wdWord9TableBehavior;
            Object autoFitBehavior = WdAutoFitBehavior.wdAutoFitWindow;

            //Добавляем таблицу и получаем объект wordtable 
            Word.Table wordtable = wordDocument.Tables.Add(wordrange, kolich + 1, 5, ref defaultTableBehavior, ref autoFitBehavior);
            Word.Table tbl = wordDocument.Tables[1];
            tbl.Range.Font.Size = 9;
            tbl.Range.Paragraphs.LineSpacing = 12;
            
            /// Вызов метода для формирования заголовка таблицы
            Zagalovok(tbl);

            int rowNumber = startPosition; 
            foreach (FileInfo fl in dir.GetFiles())
            {
                tbl.Rows[rowNumber+1].Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                tbl.Cell(rowNumber+1, 1).Range.Text = rowNumber.ToString();
                tbl.Cell(rowNumber+1, 2).Range.Text = GiveFileDiscription(fl.Name);
                tbl.Cell(rowNumber+1, 3).Range.Text = fl.Name;
                tbl.Cell(rowNumber+1, 4).Range.Text = fl.Length.ToString("N0");
                tbl.Cell(rowNumber+1, 5).Range.Text = fl.CreationTime.ToString("g");
                rowNumber++;      
            }
            /// сохраняем файл и закрываем его
            wordDocument.SaveAs2(theWay + fName);
            wordDocument.Close();
            /// делаем красиво по памяти, закрываем Word и выводим сообщение, что всё готово
            wordDocument = null;
            fileOpen.Quit();
            fileOpen = null;
            MessageBox.Show("Файл сформирован");
        }

        

        /// Метод вызова окна для уточнения - а все ли файлы есть?
        public bool WindowOfClarify(List<string> listFiles)
        {
            var result = System.Windows.MessageBox.Show(
                                                "Не хватет следующих файлов:\n\n" + String.Join("\n", listFiles) + "\n\nПродолжить?",
                                                "ВНИМАНИЕ!!!",
                                                MessageBoxButton.YesNo,
                                                MessageBoxImage.Warning);
            return result == MessageBoxResult.Yes;
        }
       

        public bool CheckFileList (string selectedPath, List<string> listOfFiles)
        {
            DirectoryInfo dir = new DirectoryInfo(selectedPath);
            List<string> listFiles = new List<string>();
            List<string> listADK = new List<string>();
            return true;
        }
    }


}
