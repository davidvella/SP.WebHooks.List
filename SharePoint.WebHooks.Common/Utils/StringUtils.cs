using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SharePoint.WebHooks.Common.Utils
{
    public class StringUtils
    {
        public static IList<string> FormatTextToSentences(string text)
        {
            // figure out, which sentence length we're using based on set accuracylevel. The default value is 5120 (set by the API)
            const int splitStringLimit = 2560;

            return Regex.Matches(text, @"(.{1," + splitStringLimit + @"})(?:\s|$)").Cast<Match>().Select(m => m.Value).ToArray();
        }
    }
}
