using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace OutlookEventsPlugin
{
    partial class Ribbon
    {
        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnPrintEvents = this.Factory.CreateRibbonButton();
            this.btnPrintDay = this.Factory.CreateRibbonButton();

            // tab1
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabCalendar";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Календарь";

            // group1
            this.group1.Items.Add(this.btnPrintEvents);
            this.group1.Items.Add(this.btnPrintDay);
            this.group1.Label = "Печать событий";

            // btnPrintEvents
            this.btnPrintEvents.Label = "Печать выбранных событий";
            this.btnPrintEvents.Name = "btnPrintEvents";
            this.btnPrintEvents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrintEvents_Click);

            // btnPrintDay
            this.btnPrintDay.Label = "Печать событий за день";
            this.btnPrintDay.Name = "btnPrintDay";
            this.btnPrintDay.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrintDay_Click);

            // Ribbon
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrintEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrintDay;
    }
} 