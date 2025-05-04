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
        private Font _titleFont;
        private Font _contentFont;
        private Font _durationFont;
        private int _currentPage;
        private int _totalPages;
        private List<string> _pages;
        private const int MARGIN = 50;
        private const int LINES_PER_PAGE = 50;
        private const int LINE_HEIGHT = 20;
        private const int DIVIDER_LINE_HEIGHT = 2;

        public CalendarPrintTemplate(Microsoft.Office.Interop.Outlook.Application outlookApp)
        {
            _outlookApp = outlookApp;
            _titleFont = new Font("Arial", 10, FontStyle.Bold);
            _contentFont = new Font("Arial", 10);
            _durationFont = new Font("Arial", 10, FontStyle.Bold);
        }

        public void PrintSelectedAppointments()
        {
            try
            {
                var explorer = _outlookApp.ActiveExplorer();
                if (explorer == null) return;

                var selection = explorer.Selection;
                if (selection == null || selection.Count == 0) return;

                var appointments = new List<AppointmentItem>();
                foreach (var item in selection)
                {
                    if (item is AppointmentItem appointment)
                    {
                        appointments.Add(appointment);
                    }
                }

                if (appointments.Count == 0) return;

                _printContent = ParseAppointments(appointments);
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

        private string ParseAppointments(List<AppointmentItem> appointments)
        {
            var printDocument = new StringBuilder();
            foreach (var appointment in appointments)
            {
                var details = new StringBuilder();
                details.AppendLine($"Тема: {appointment.Subject}");
                
                var duration = appointment.End - appointment.Start;
                var days = duration.Days;
                var hours = duration.Hours;
                var minutes = duration.Minutes;
                var durationParts = new List<string>();
                if (days > 0) durationParts.Add($"{days}д");
                if (hours > 0) durationParts.Add($"{hours}ч");
                if (minutes > 0) durationParts.Add($"{minutes}м");
                var durationText = string.Join(" ", durationParts);
                details.AppendLine($"Время: {appointment.Start.ToString("dd.MM.yyyy")}, {appointment.Start.ToString("HH:mm")} - {appointment.End.ToString("HH:mm")} ({durationText})");

                if (appointment.Recipients.Count > 0)
                {
                    var recipients = string.Join("; ", appointment.Recipients.Cast<Recipient>().Select(r => r.Name));
                    details.AppendLine($"Участники: {recipients}");
                }
                details.AppendLine("DIVIDER");
                
                printDocument.Append(details.ToString());
            }
            return printDocument.ToString();
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
                var stringFormat = new StringFormat();
                stringFormat.Trimming = StringTrimming.Word;
                
                // Определяем ширину колонок в процентах
                const float LABEL_COLUMN_PERCENT = 0.2f; // 20% ширины страницы
                const float VALUE_COLUMN_PERCENT = 0.8f; // 80% ширины страницы
                
                // Вычисляем ширину колонок
                var availableWidth = e.PageBounds.Width - 2 * MARGIN;
                var labelColumnWidth = availableWidth * LABEL_COLUMN_PERCENT;
                var valueColumnWidth = availableWidth * VALUE_COLUMN_PERCENT;
                
                foreach (var line in lines)
                {
                    if (line == "DIVIDER")
                    {
                        e.Graphics.DrawLine(new Pen(Color.Black, DIVIDER_LINE_HEIGHT), 
                            MARGIN, y, e.PageBounds.Width - MARGIN, y);
                        y += LINE_HEIGHT;
                    }
                    else if (!string.IsNullOrWhiteSpace(line))
                    {
                        var parts = line.Split(new[] { ':' }, 2);
                        if (parts.Length == 2)
                        {
                            // Отрисовка метки
                            var labelRect = new RectangleF(MARGIN, y, labelColumnWidth, e.PageBounds.Height - y);
                            e.Graphics.DrawString(parts[0] + ":", _titleFont, Brushes.Black, labelRect, stringFormat);
                            
                            // Отрисовка значения
                            var valueRect = new RectangleF(MARGIN + labelColumnWidth, y, valueColumnWidth, e.PageBounds.Height - y);
                            var valueText = parts[1]?.Trim() ?? string.Empty;
                            
                            // Проверяем, содержит ли текст длительность в скобках
                            if (valueText.Contains("(") && valueText.Contains(")") && valueText.Contains("ч") || valueText.Contains("д") || valueText.Contains("м"))
                            {
                                var startIndex = valueText.LastIndexOf("(");
                                var endIndex = valueText.LastIndexOf(")");
                                
                                if (startIndex >= 0 && endIndex > startIndex)
                                {
                                    var beforeDuration = valueText.Substring(0, startIndex);
                                    var duration = valueText.Substring(startIndex, endIndex - startIndex + 1);
                                    
                                    // Отрисовка текста до длительности
                                    e.Graphics.DrawString(beforeDuration, _contentFont, Brushes.Black, valueRect, stringFormat);
                                    
                                    // Отрисовка длительности жирным шрифтом
                                    var durationWidth = e.Graphics.MeasureString(beforeDuration, _contentFont, (int)valueColumnWidth, stringFormat).Width;
                                    var durationRect = new RectangleF(MARGIN + labelColumnWidth + durationWidth, y, valueColumnWidth - durationWidth, e.PageBounds.Height - y);
                                    e.Graphics.DrawString(duration, _durationFont, Brushes.Black, durationRect, stringFormat);
                                }
                                else
                                {
                                    e.Graphics.DrawString(valueText, _contentFont, Brushes.Black, valueRect, stringFormat);
                                }
                            }
                            else
                            {
                                e.Graphics.DrawString(valueText, _contentFont, Brushes.Black, valueRect, stringFormat);
                            }
                            
                            // Вычисляем высоту для следующей строки
                            var labelHeight = e.Graphics.MeasureString(parts[0] + ":", _titleFont, (int)labelColumnWidth, stringFormat).Height;
                            var valueHeight = e.Graphics.MeasureString(parts[1].Trim(), _contentFont, (int)valueColumnWidth, stringFormat).Height;
                            y += (int)Math.Max(labelHeight, valueHeight);
                        }
                        else
                        {
                            // Для строк без разделителя (например, "Участники:")
                            var layoutRect = new RectangleF(MARGIN, y, e.PageBounds.Width - 2 * MARGIN, e.PageBounds.Height - y);
                            e.Graphics.DrawString(line, _titleFont, Brushes.Black, layoutRect, stringFormat);
                            y += (int)e.Graphics.MeasureString(line, _titleFont, (int)(e.PageBounds.Width - 2 * MARGIN), stringFormat).Height;
                        }
                    }
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