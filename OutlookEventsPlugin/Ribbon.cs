using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace OutlookEventsPlugin
{
    public partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void btnPrintEvents_Click(object sender, RibbonControlEventArgs e)
        {
            var printTemplate = new CalendarPrintTemplate(Globals.ThisAddIn.Application);
            printTemplate.PrintSelectedAppointments();
        }
    }
} 