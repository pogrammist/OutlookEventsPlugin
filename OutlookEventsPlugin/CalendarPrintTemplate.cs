using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Printing;

namespace OutlookEventsPlugin
{
    public class CalendarPrintTemplate
    {
        private readonly Microsoft.Office.Interop.Outlook.Application _outlookApp;
        private string _printContent;
        private Font _printFont;
        private int _currentPage;
        private int _totalPages;
        private List<string> _pages;
        private const int MARGIN = 50;
        private const int LINES_PER_PAGE = 50;

        public CalendarPrintTemplate(Microsoft.Office.Interop.Outlook.Application outlookApp)
        {
            _outlookApp = outlookApp;
            _printFont = new Font("Arial", 10);
        }

        public void PrintSelectedEvents()
        {
            try
            {
                var explorer = _outlookApp.ActiveExplorer();
                if (explorer == null) return;

                var selection = explorer.Selection;
                if (selection == null || selection.Count == 0) return;

                var events = new List<AppointmentItem>();
                foreach (var item in selection)
                {
                    if (item is AppointmentItem appointment)
                    {
                        events.Add(appointment);
                    }
                }

                if (events.Count == 0) return;

                var printDocument = new StringBuilder();
                printDocument.AppendLine("Детальная информация о событиях календаря");
                printDocument.AppendLine("=======================================");
                printDocument.AppendLine();

                foreach (var appointment in events)
                {
                    printDocument.AppendLine($"Событие: {appointment.Subject}");
                    printDocument.AppendLine($"Начало: {appointment.Start}");
                    printDocument.AppendLine($"Окончание: {appointment.End}");
                    printDocument.AppendLine($"Место: {appointment.Location}");
                    printDocument.AppendLine($"Описание: {appointment.Body}");
                    printDocument.AppendLine();

                    if (appointment.Recipients.Count > 0)
                    {
                        printDocument.AppendLine("Участники:");
                        foreach (Recipient recipient in appointment.Recipients)
                        {
                            printDocument.AppendLine($"- {recipient.Name} ({recipient.Address})");
                        }
                    }
                    printDocument.AppendLine("---------------------------------------");
                    printDocument.AppendLine();
                }

                _printContent = printDocument.ToString();
                _currentPage = 0;

                // Разбиваем текст на страницы
                _pages = SplitTextIntoPages(_printContent);

                // Создаем документ для печати
                var doc = new PrintDocument();
                doc.PrintPage += Doc_PrintPage;
                doc.EndPrint += Doc_EndPrint;

                // Открываем диалог печати
                var printDialog = new PrintDialog();
                printDialog.Document = doc;
                if (printDialog.ShowDialog() == DialogResult.OK)
                {
                    doc.Print();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Ошибка при печати: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<string> SplitTextIntoPages(string text)
        {
            var pages = new List<string>();
            var lines = text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            var currentPage = new StringBuilder();
            var lineCount = 0;

            foreach (var line in lines)
            {
                currentPage.AppendLine(line);
                lineCount++;

                if (lineCount >= LINES_PER_PAGE)
                {
                    pages.Add(currentPage.ToString());
                    currentPage.Clear();
                    lineCount = 0;
                }
            }

            if (currentPage.Length > 0)
            {
                pages.Add(currentPage.ToString());
            }

            _totalPages = pages.Count;
            return pages;
        }

        private void Doc_PrintPage(object sender, PrintPageEventArgs e)
        {
            if (_currentPage < _pages.Count)
            {
                var y = MARGIN;
                var lines = _pages[_currentPage].Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                
                foreach (var line in lines)
                {
                    e.Graphics.DrawString(line, _printFont, Brushes.Black, MARGIN, y);
                    y += (int)_printFont.GetHeight(e.Graphics);
                }

                _currentPage++;
                e.HasMorePages = _currentPage < _pages.Count;
            }
        }

        private void Doc_EndPrint(object sender, PrintEventArgs e)
        {
            _currentPage = 0;
            _pages.Clear();
        }
    }
} 