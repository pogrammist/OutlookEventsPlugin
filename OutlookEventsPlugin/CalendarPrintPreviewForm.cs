using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace OutlookEventsPlugin
{
    public class CalendarPrintPreviewForm : Form
    {
        private PrintDocument _printDocument;
        private PrintPreviewControl _previewControl;

        public CalendarPrintPreviewForm(PrintDocument printDocument)
        {
            _printDocument = printDocument;
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "Предпросмотр печати";
            this.Size = new Size(1200, Screen.PrimaryScreen.WorkingArea.Height);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Icon = SystemIcons.Information;
            this.ShowIcon = true;

            _previewControl = new PrintPreviewControl
            {
                Dock = DockStyle.Fill,
                Document = _printDocument,
                Zoom = 1.0
            };

            var buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 40
            };

            var printButton = new Button
            {
                Text = "Печать",
                Location = new Point(10, 10),
                Size = new Size(100, 25)
            };
            printButton.Click += (s, e) =>
            {
                var printDialog = new PrintPreviewDialog
                {
                    Document = _printDocument
                };

                if (printDialog.ShowDialog() == DialogResult.OK)
                {
                    _printDocument.Print();
                    this.Close();
                }
            };

            var closeButton = new Button
            {
                Text = "Закрыть",
                Location = new Point(120, 10),
                Size = new Size(100, 25)
            };
            closeButton.Click += (s, e) => this.Close();

            buttonPanel.Controls.Add(printButton);
            buttonPanel.Controls.Add(closeButton);

            this.Controls.Add(_previewControl);
            this.Controls.Add(buttonPanel);
        }
    }
}