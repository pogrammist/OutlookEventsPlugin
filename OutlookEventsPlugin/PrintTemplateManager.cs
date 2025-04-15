using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace OutlookEventsPlugin
{
    [ComVisible(true)]
    public class PrintTemplateManager : IRibbonExtensibility
    {
        private readonly Application _outlookApp;

        public PrintTemplateManager(Application outlookApp)
        {
            _outlookApp = outlookApp;
        }

        public string GetCustomUI(string ribbonID)
        {
            return @"
                <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
                    <backstage>
                        <button id='customPrintTemplate' 
                                label='Детальная печать событий' 
                                onAction='OnCustomPrintTemplate'
                                imageMso='PrintPreview' />
                    </backstage>
                </customUI>";
        }

        public void OnCustomPrintTemplate(IRibbonControl control)
        {
            try
            {
                var printTemplate = new CalendarPrintTemplate(_outlookApp);
                printTemplate.PrintSelectedEvents();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Ошибка при печати: {ex.Message}", 
                    "Ошибка", 
                    System.Windows.Forms.MessageBoxButtons.OK, 
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
} 