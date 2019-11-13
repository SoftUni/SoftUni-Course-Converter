using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;

namespace SoftUniConverterCommon
{
    public enum Language { EN, BG };

    public class ConverterUtils
    {
        public static object GetObjectProperty(object obj, string propName)
        {
            object prop = obj.GetType().InvokeMember(
                "Item", BindingFlags.Default | BindingFlags.GetProperty,
                null, obj, new object[] { propName });
            object propValue = prop.GetType().InvokeMember(
                "Value", BindingFlags.Default | BindingFlags.GetProperty,
                null, prop, new object[] { });
            return propValue;
        }

        public static void SetObjectProperty(object obj, string propName, object propValue)
        {
            object prop = obj.GetType().InvokeMember(
                "Item", BindingFlags.Default | BindingFlags.GetProperty,
                null, obj, new object[] { propName });
            prop.GetType().InvokeMember(
                "Value", BindingFlags.Default | BindingFlags.SetProperty,
                null, prop, new object[] { propValue });
        }

        private static readonly HashSet<string> EnglishTitleCaseIgnoredWords = new HashSet<string> {
            "a", "an", "the", "is", "vs",
            "and", "or", "in", "of", "by", "from", "at", "off", "to",
            "into", "about", "onto", "for", "with"
        };

        public static string FixEnglishTitleCharacterCasing(string text)
        {
            string EnglishWordToTitleCase(string word)
            {
                if (string.IsNullOrEmpty(word))
                    return word;

                // Handle normal words like "program"
                if (char.ToLower(word[0]) >= 'a' && char.ToLower(word[0]) <= 'z')
                    return "" + char.ToUpper(word[0]) + word.Substring(1);

                // Handle words like "[Run]" or "(maybe)"
                if (word.Length > 1 && char.ToLower(word[1]) >= 'a' && char.ToLower(word[1]) <= 'z')
                    return "" + word[0] + char.ToUpper(word[1]) + word.Substring(2);

                return word;
            }

            if (string.IsNullOrEmpty(text))
                return text;

            text = text.Replace(" - ", " – ");

            string[] words = text.Split(' ');
            if (words.Length > 0)
            {
                // Always start with capital letter
                words[0] = EnglishWordToTitleCase(words[0]);
            }
            for (int i = 1; i < words.Length; i++)
            {
                string wordOnly = words[i].Trim(' ', ',', ';', '?', '!', '.', '(', ')');
                if (EnglishTitleCaseIgnoredWords.Contains(wordOnly.ToLower()))
                {
                    // Special word (like preposition / conjunctions) -> 
                    // lowercase it (unless it is ALL CAPS)
                    if (wordOnly != wordOnly.ToUpper())
                        words[i] = words[i].ToLower();
                }
                else
                {
                    // Normal word (non-special) -> capitalize its first letter
                    words[i] = EnglishWordToTitleCase(words[i]);
                }
            }

            string result = string.Join(" ", words);
            return result;
        }

        public static string TruncateString(string str, int maxLength)
        {
            if (str == null)
                return "";
            str = str.Trim();
            if (str.Length > maxLength)
                str = str.Substring(0, maxLength) + "...";
            return str;
        }

        public static bool KillAllProcesses(string processName)
        {
            Process[] processes = Process.GetProcessesByName(processName);
            int killedProcessesCount = 0;
            foreach (Process process in processes)
            {
                try
                {
                    process.Kill();
                    killedProcessesCount++;
                }
                catch
                {
                    // Ignore the exception: the process cannot be killed for some reason
                }
            }
            return (killedProcessesCount > 0);
        }
    }
}
