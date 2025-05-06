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
        private PrintContext _printContext;

        public CalendarPrintTemplate(Microsoft.Office.Interop.Outlook.Application outlookApp)
        {
            _outlookApp = outlookApp;
            _printContext = new PrintContext();
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

                ShowPrintPreview(appointments);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Ошибка при печати: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowPrintPreview(List<AppointmentItem> appointments)
        {
            _printContext.Content = ParseAppointments(appointments);
            _printContext.CurrentPage = 0;

            // Разбиваем текст на страницы
            _printContext.Pages = SplitTextIntoPages(_printContext.Content);

            // Создаем документ для печати
            var doc = new PrintDocument();
            doc.PrintPage += Doc_PrintPage;
            doc.EndPrint += Doc_EndPrint;

            var previewForm = new CalendarPrintPreviewDialog(doc, _printContext);
            previewForm.ShowDialog();
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

                if (lineCount >= 60) // LINES_PER_PAGE
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

            _printContext.TotalPages = pages.Count;
            return pages;
        }

        private void Doc_PrintPage(object sender, PrintPageEventArgs e)
        {
            if (_printContext.CurrentPage < _printContext.Pages.Count)
            {
                var y = PrintContext.MARGIN;
                var lines = _printContext.Pages[_printContext.CurrentPage].Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                var stringFormat = new StringFormat();
                stringFormat.Trimming = StringTrimming.Word;
                
                // Определяем ширину колонок в процентах
                const float LABEL_COLUMN_PERCENT = 0.2f; // 20% ширины страницы
                const float VALUE_COLUMN_PERCENT = 0.8f; // 80% ширины страницы
                
                // Вычисляем ширину колонок
                var availableWidth = e.PageBounds.Width - 2 * PrintContext.MARGIN;
                var labelColumnWidth = availableWidth * LABEL_COLUMN_PERCENT;
                var valueColumnWidth = availableWidth * VALUE_COLUMN_PERCENT;
                
                foreach (var line in lines)
                {
                    if (line == "DIVIDER")
                    {
                        e.Graphics.DrawLine(new Pen(Color.Gray, PrintContext.DIVIDER_LINE_HEIGHT), 
                            PrintContext.MARGIN, y, e.PageBounds.Width - PrintContext.MARGIN, y);
                        y += PrintContext.LINE_HEIGHT;
                    }
                    else if (!string.IsNullOrWhiteSpace(line))
                    {
                        var parts = line.Split(new[] { ':' }, 2);
                        if (parts.Length == 2)
                        {
                            // Отрисовка метки
                            var labelRect = new RectangleF(PrintContext.MARGIN, y, labelColumnWidth, e.PageBounds.Height - y);
                            e.Graphics.DrawString(parts[0] + ":", _printContext.TitleFont, Brushes.Black, labelRect, stringFormat);
                            
                            // Отрисовка значения
                            var valueRect = new RectangleF(PrintContext.MARGIN + labelColumnWidth, y, valueColumnWidth, e.PageBounds.Height - y);
                            var valueText = parts[1]?.Trim() ?? string.Empty;
                            
                            // Проверяем, содержит ли текст длительность в скобках
                            if (valueText.Contains("(") && valueText.Contains(")") && (valueText.Contains("ч") || valueText.Contains("д") || valueText.Contains("м")))
                            {
                                var startIndex = valueText.LastIndexOf("(");
                                var endIndex = valueText.LastIndexOf(")");
                                
                                if (startIndex >= 0 && endIndex > startIndex)
                                {
                                    var beforeDuration = valueText.Substring(0, startIndex);
                                    var duration = valueText.Substring(startIndex, endIndex - startIndex + 1);
                                    
                                    // Отрисовка текста до длительности
                                    e.Graphics.DrawString(beforeDuration, _printContext.ContentFont, Brushes.Black, valueRect, stringFormat);
                                    
                                    // Отрисовка длительности жирным шрифтом
                                    var durationWidth = e.Graphics.MeasureString(beforeDuration, _printContext.ContentFont, (int)valueColumnWidth, stringFormat).Width;
                                    var durationRect = new RectangleF(PrintContext.MARGIN + labelColumnWidth + durationWidth, y, valueColumnWidth - durationWidth, e.PageBounds.Height - y);
                                    e.Graphics.DrawString(duration, _printContext.DurationFont, Brushes.Black, durationRect, stringFormat);
                                }
                                else
                                {
                                    e.Graphics.DrawString(valueText, _printContext.ContentFont, Brushes.Black, valueRect, stringFormat);
                                }
                            }
                            else
                            {
                                e.Graphics.DrawString(valueText, _printContext.ContentFont, Brushes.Black, valueRect, stringFormat);
                            }
                            
                            // Вычисляем высоту для следующей строки
                            var labelHeight = e.Graphics.MeasureString(parts[0] + ":", _printContext.TitleFont, (int)labelColumnWidth, stringFormat).Height;
                            var valueHeight = e.Graphics.MeasureString(parts[1].Trim(), _printContext.ContentFont, (int)valueColumnWidth, stringFormat).Height;
                            y += (int)Math.Max(labelHeight, valueHeight);
                        }
                        else
                        {
                            // Для строк без разделителя (например, "Участники:")
                            var layoutRect = new RectangleF(PrintContext.MARGIN, y, e.PageBounds.Width - 2 * PrintContext.MARGIN, e.PageBounds.Height - y);
                            e.Graphics.DrawString(line, _printContext.TitleFont, Brushes.Black, layoutRect, stringFormat);
                            y += (int)e.Graphics.MeasureString(line, _printContext.TitleFont, (int)(e.PageBounds.Width - 2 * PrintContext.MARGIN), stringFormat).Height;
                        }
                    }
                }

                _printContext.CurrentPage++;
                e.HasMorePages = _printContext.CurrentPage < _printContext.Pages.Count;
            }
        }

        private void Doc_EndPrint(object sender, PrintEventArgs e)
        {
            _printContext.CurrentPage = 0;
        }
    }
} 