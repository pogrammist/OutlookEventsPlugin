using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;

namespace OutlookEventsPlugin
{
    public class PrintContext
    {
        public string Content { get; set; }
        public int CurrentPage { get; set; }
        public int TotalPages { get; set; }
        public List<string> Pages { get; set; }
        public Font TitleFont { get; set; }
        public Font ContentFont { get; set; }
        public Font DurationFont { get; set; }
        public const int MARGIN = 50;
        public const int LINE_HEIGHT = 20;
        public const int DIVIDER_LINE_HEIGHT = 2;

        public PrintContext()
        {
            TitleFont = new Font("Arial", 10, FontStyle.Bold);
            ContentFont = new Font("Arial", 10);
            DurationFont = new Font("Arial", 10, FontStyle.Bold);
            Pages = new List<string>();
            CurrentPage = 0;
        }

        public void Dispose()
        {
            TitleFont?.Dispose();
            ContentFont?.Dispose();
            DurationFont?.Dispose();
        }
    }
} 