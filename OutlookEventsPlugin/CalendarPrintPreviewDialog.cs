using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace OutlookEventsPlugin
{
    public class CalendarPrintPreviewDialog : Form
    {
        private PrintDocument _printDocument;
        private PrintPreviewControl _previewControl;
        private PrintContext _printContext;
        private Label _pageInfoLabel;
        private Button _prevPageButton;
        private Button _nextPageButton;

        public CalendarPrintPreviewDialog(PrintDocument printDocument, PrintContext printContext)
        {
            _printDocument = printDocument;
            _printContext = printContext;
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
                Zoom = 1.0,
                StartPage = 0,
                Rows = 1,
                Columns = 1,
                UseAntiAlias = true
            };

            var buttonPanel = new Panel
            {
                Dock = DockStyle.Top,
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
                var printDialog = new PrintDialog
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

            _prevPageButton = new Button
            {
                Text = "←",
                Location = new Point(230, 10),
                Size = new Size(30, 25),
                Enabled = false
            };
            _prevPageButton.Click += (s, e) =>
            {
                if (_previewControl.StartPage > 0)
                {
                    _previewControl.StartPage--;
                    UpdatePageInfo();
                }
            };

            _nextPageButton = new Button
            {
                Text = "→",
                Location = new Point(270, 10),
                Size = new Size(30, 25),
                Enabled = false
            };
            _nextPageButton.Click += (s, e) =>
            {
                if (_previewControl.StartPage < _printContext.TotalPages - 1)
                {
                    _previewControl.StartPage++;
                    UpdatePageInfo();
                }
            };

            _pageInfoLabel = new Label
            {
                Location = new Point(310, 15),
                Size = new Size(200, 20),
                Text = $"Страница 1 из {_printContext.TotalPages}"
            };

            buttonPanel.Controls.Add(printButton);
            buttonPanel.Controls.Add(closeButton);
            buttonPanel.Controls.Add(_prevPageButton);
            buttonPanel.Controls.Add(_nextPageButton);
            buttonPanel.Controls.Add(_pageInfoLabel);

            this.Controls.Add(_previewControl);
            this.Controls.Add(buttonPanel);

            UpdatePageInfo();
        }

        private void UpdatePageInfo()
        {
            _pageInfoLabel.Text = $"Страница {_previewControl.StartPage + 1} из {_printContext.TotalPages}";
            _prevPageButton.Enabled = _previewControl.StartPage > 0;
            _nextPageButton.Enabled = _previewControl.StartPage < _printContext.TotalPages - 1;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            _printContext.Dispose();
        }
    }
} 