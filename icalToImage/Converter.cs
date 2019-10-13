using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using ical = Ical.Net;
using CoreHtmlToImage;
using System.Net;
using System.Globalization;
using System.IO;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace IcalToImage
{
    /// <summary>
    /// Main class to convert a calendar to an image.
    /// </summary>
    public class Converter
    {
        #region Private Fields

        private readonly string Footer =
@"    </table>
</body>
</html>";

        private readonly string Header1 =
@"<!DOCTYPE HTML>
<html>
<head>
    <meta charset=""UTF-8"">
    <style>
";

        private readonly string Header2 =
@"    </style>
</head>
<body>
    <table>
";

        #endregion Private Fields

        #region Public Constructors

        /// <summary>
        /// Constructor.
        /// </summary>
        public Converter()
        {
            if (Client == null)
                Client = new WebClient();
            Calendar = null;
            Culture = CultureInfo.CurrentCulture;
            CSS = new Dictionary<Tag, string>();
            CSS[Tag.TABLE] =
@"border: 2px solid black;
border-collapse: collapse;
empty-cells: show;
font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande', 'Lucida Sans', Arial, sans-serif;";
            CSS[Tag.HEAD] =
@"border-bottom-style: double;";
            CSS[Tag.CELLS] =
@"max-width: 160px;
border: 1px solid black;
text-align: center;
padding: 5px;";
            CSS[Tag.EMPTY_CELLS] =
@"border: none;
border-right: 1px solid black;
background-color: lightgrey;";
            CSS[Tag.HEAD_CELLS] =
@"font-weight: bold;";
            CSS[Tag.ROWS] =
@"flex: 1;";
            CSS[Tag.BODY] =
@"display: flex;";
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="calendar">calendar to convert</param>
        public Converter(ical.Calendar calendar) : this()
        {
            Calendar = calendar;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="uri">Uri from which to get the calendar</param>
        public Converter(Uri uri) : this() => Calendar = ical.Calendar.Load(new MemoryStream(Client.DownloadData(uri)));

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="icalendar">String containing the data</param>
        public Converter(string icalendar) : this() => Calendar = ical.Calendar.Load(icalendar);

        #endregion Public Constructors

        #region Public Enums

        public enum Tag
        {
            TABLE,
            CELLS,
            HEAD_CELLS,
            HEAD,
            BODY,
            EMPTY_CELLS,
            ROWS
        }

        #endregion Public Enums

        #region Public Properties

        /// <summary>
        /// Calendar to convert
        /// </summary>
        public ical.Calendar Calendar { get; set; }

        /// <summary>
        /// Culture used to generate the image.
        /// </summary>
        public CultureInfo Culture { get; set; }

        #endregion Public Properties

        #region Private Properties

        private static WebClient Client { get; set; }
        private Dictionary<Tag, string> CSS { get; set; }

        #endregion Private Properties

        #region Public Methods

        /// <summary>
        /// Change the css to use for the generated image. (The calendar is first covnerted into html).
        /// </summary>
        /// <param name="tag"></param>
        /// <param name="css"></param>
        public void ChangeCSS(Tag tag, string css) => CSS[tag] = css;

        /// <summary>
        /// Generate the image from the calendar.
        /// </summary>
        /// <param name="width">Width of the image</param>
        /// <param name="firstDay">first day of the range</param>
        /// <param name="lastDay">last day of the range</param>
        /// <returns>image</returns>
        public Bitmap ConvertRangeToBitmap(uint width, DateTime firstDay, DateTime lastDay)
        {
            firstDay = firstDay.Date;
            lastDay = lastDay.Date;
            var min = firstDay > lastDay ? lastDay : firstDay;
            var max = firstDay < lastDay ? lastDay : firstDay;
            var curr = min;
            var days = new List<DateTime>();
            do
                days.Add(curr);
            while (curr < max);
            return ConvertToBitmap(width, days.ToArray());
        }

        /// <summary>
        /// Generate the image from the calendar.
        /// </summary>
        /// <param name="width">Width of the image</param>
        /// <param name="days">days to display</param>
        /// <exception cref="NullReferenceException">days is null</exception>
        /// <returns>image</returns>
        public Bitmap ConvertToBitmap(uint width, params DateTime[] days)
        {
            int minquarter = 24 * 4;
            int maxquarter = 0;
            var events = new Dictionary<DateTime, List<Ev>>();
            foreach (var d in days)
                events.Add(d.Date, new List<Ev>());
            foreach (var ev in Calendar.Events)
            {
                if (events.ContainsKey(ev.Start.Date))
                    events[ev.Start.Date].Add(new Ev()
                    {
                        description = ev.Summary,
                        duration = (int)((ev.End.Value - ev.Start.Value).TotalMinutes / 15),
                        start = ev.Start.Value.Hour * 4 + ev.Start.Value.Minute / 15,
                        actualStart = ev.Start.Value,
                        actualDuration = ev.Duration
                    });
            }
            foreach (var d in events)
            {
                d.Value.Sort((l, r) => l.actualStart.CompareTo(r.actualStart));
                foreach (var ev in d.Value)
                {
                    if (ev.start < minquarter)
                        minquarter = ev.start;
                    if (ev.start + ev.duration > maxquarter)
                        maxquarter = ev.start + ev.duration;
                }
            }
            int minHour = (int)Math.Floor((minquarter / 4.0));
            int maxHour = (int)Math.Ceiling((maxquarter / 4.0));

            var sb = new StringBuilder();
            sb.Append(Header1);
            foreach (var entry in CSS)
            {
                switch (entry.Key)
                {
                    case Tag.CELLS:
                        sb.Append("td, th{");
                        sb.Append(entry.Value);
                        sb.Append("}\n");
                        break;

                    case Tag.EMPTY_CELLS:
                        sb.Append("td:empty{");
                        sb.Append(entry.Value);
                        sb.Append("}\n");
                        break;

                    case Tag.HEAD:
                        sb.Append("thead{");
                        sb.Append(entry.Value);
                        sb.Append("}\n");
                        break;

                    case Tag.HEAD_CELLS:
                        sb.Append("th{");
                        sb.Append(entry.Value);
                        sb.Append("}\n");
                        break;

                    case Tag.TABLE:
                        sb.Append("table{");
                        sb.Append(entry.Value);
                        sb.Append("}\n");
                        break;

                    case Tag.ROWS:
                        sb.Append("tr{");
                        sb.Append(entry.Value);
                        sb.Append("}\n");
                        break;

                    case Tag.BODY:
                        sb.Append("tbody{");
                        sb.Append(entry.Value);
                        sb.Append("}\n");
                        break;
                }
            }
            sb.Append(Header2);
            sb.Append("<thead>\n<tr>\n<th></th>\n");
            foreach (var ev in events)
            {
                sb.Append("<th>");
                sb.Append(Culture.DateTimeFormat.GetDayName(ev.Key.DayOfWeek));
                sb.Append(" ");
                sb.Append(ev.Key.ToString("d/M", Culture.DateTimeFormat));
                sb.Append("</th>\n");
            }
            sb.Append("</tr>\n</thead>\n<tbody>\n");
            int currentTime = minHour * 4;
            var occupiedPlace = new Dictionary<DateTime, int>();
            foreach (var ev in events)
                occupiedPlace.Add(ev.Key, 0);
            for (; currentTime <= maxHour * 4; currentTime++)
            {
                sb.Append("<tr>\n");
                if (currentTime % 4 == 0)
                {
                    sb.Append("<th rowspan=\"4\">");
                    sb.Append(DateTime.Today.AddMinutes(currentTime * 15).ToString("t", Culture.DateTimeFormat));
                    sb.Append("</th>\n");
                }
                foreach (var ev in events)
                {
                    var currentEv = ev.Value.FirstOrDefault(e => e.start == currentTime);
                    if (currentEv.CompareTo(default) != 0)
                    {
                        sb.Append("<td rowspan=\"" + currentEv.duration + "\">");
                        sb.Append(currentEv.actualStart.ToString("t", Culture.DateTimeFormat));
                        sb.Append(" - ");
                        sb.Append((currentEv.actualStart + currentEv.actualDuration).ToString("t", Culture.DateTimeFormat));
                        sb.Append("<br />");
                        sb.Append(currentEv.description);
                        sb.Append("</td>\n");
                        occupiedPlace[ev.Key] = currentEv.start + currentEv.duration;
                    }
                    else if (occupiedPlace[ev.Key] <= currentTime)
                        sb.Append("<td></td>");
                }
                sb.Append("</tr>\n");
            }
            sb.Append("</tbody>");
            sb.Append(Footer);

            Bitmap img;
            using (var stream = new MemoryStream(new HtmlConverter().FromHtmlString(sb.ToString(), (int)width, ImageFormat.Png)))
                img = new Bitmap(stream);
            return img;
        }

        #endregion Public Methods

        #region Private Structs

        private struct Ev : IComparable<Ev>
        {
            #region Public Fields

            public TimeSpan actualDuration;
            public DateTime actualStart;
            public string description;
            public int duration;
            public int start;

            #endregion Public Fields

            #region Public Methods

            public int CompareTo([AllowNull] Ev other) => (start * 4 + duration).CompareTo(other.start * 4 + other.duration);

            #endregion Public Methods
        }

        #endregion Private Structs
    }
}