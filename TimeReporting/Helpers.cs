using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace OutlookAddIn1
{
    class Helpers
    {
        private static string ParseString(string inputData, string pattern)
        {
            RegexOptions regexOptions = RegexOptions.None;
            Regex regex = new Regex(pattern, regexOptions);
            foreach (Match match in regex.Matches(inputData))
            {
                if (match.Success)
                {
                    return match.Groups[1].Value;                    
                }
            }
            return "";
        }

        public static int ParseID(string inputData)
        {
            string parsed = ParseString(inputData, @"^\[tfs([0-9]+)\].*$");
            if (parsed != "")
                return Convert.ToInt32(parsed);
            return 0;
        }

        public static string ParseTitle(string inputData)
        {
            return ParseString(inputData, @"^\[tfs[0-9]+\](.*)$").Trim();
        }

        public static string ParseTitleWithTime(string inputData)
        {
            return ParseString(inputData, @"^\[tfs[0-9]+](.*)\*.*\*$").Trim();
        }

        public static string GenerateSubject(int id, string title, bool time = false)
        {
            string subject = "[tfs" + id + "]";
            if (title != "")
                subject += " " + title;
            if (subject.Length > 220)
                subject = subject.Substring(0, 220);
            if (time)
                subject += " *" + DateTime.Now.ToString() + "*";
                
            return subject;
        }

        public static void DebugInfo(string info)
        {
            if (Settings.debugFile != "")
            {
                try
                {
                    var file = File.AppendText(Settings.debugFile);
                    file.WriteLine(DateTime.Now.ToString() + ": " + info);
                    file.Close();
                }
                catch (Exception)
                { }
            }                
        }

        public static int GetIDFromClipboard()
        {
            try
            {
                string parsed = Helpers.ParseString(Clipboard.GetText(), @"^Task ([0-9]+) : .*$");
                if (parsed != "")
                    return Convert.ToInt32(parsed);
                parsed = Helpers.ParseString(Clipboard.GetText(), @"^Task ([0-9]+):.*$");
                if (parsed != "")
                    return Convert.ToInt32(parsed);
                parsed = Helpers.ParseString(Clipboard.GetText(), @"^([0-9]+)$");
                if (parsed != "")
                    return Convert.ToInt32(parsed);
            }
            catch (Exception)
            { }
            return 0;
        }

    }
}
