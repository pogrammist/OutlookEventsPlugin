using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace OutlookEventsPlugin
{
    public class CalendarPrintTemplate
    {
        private readonly Microsoft.Office.Interop.Outlook.Application _outlookApp;

        public CalendarPrintTemplate(Microsoft.Office.Interop.Outlook.Application outlookApp)
        {
            _outlookApp = outlookApp;
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

                // Открываем диалог печати
                var printDialog = new PrintDialog();
                if (printDialog.ShowDialog() == DialogResult.OK)
                {
                    // Здесь можно добавить код для отправки на печать
                    MessageBox.Show("Документ отправлен на печать", "Печать", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Ошибка при печати: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
} 