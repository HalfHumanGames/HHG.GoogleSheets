using System.Text;
using System.Text.RegularExpressions;
using UnityEditor;

namespace HHG.GoogleSheets.Runtime
{
    public enum Case
    {
        None,
        Pascal,
        Title,
        Snake,
        Nicified,
    }

    public static class CaseExtensions
    {
        public static string ToCase(this Case casing, string input)
        {
            switch (casing)
            {
                case Case.None:
                    return input;

                case Case.Pascal:
                    return ToPascalCase(input);

                case Case.Title:
                    return ToTitleCase(input);

                case Case.Snake:
                    return ToSnakeCase(input);

                case Case.Nicified:
                    return ObjectNames.NicifyVariableName(input);

                default:
                    return input;
            }
        }

        private static string ToPascalCase(string input)
        {
            string[] words = SplitWords(input);
            StringBuilder sb = new StringBuilder();

            foreach (string word in words)
            {
                if (word.Length > 0)
                {
                    sb.Append(char.ToUpperInvariant(word[0]));
                    if (word.Length > 1)
                    {
                        sb.Append(word.Substring(1).ToLowerInvariant());
                    }
                }
            }

            return sb.ToString();
        }

        private static string ToTitleCase(string input)
        {
            string[] words = SplitWords(input);
            for (int i = 0; i < words.Length; i++)
            {
                if (words[i].Length > 0)
                {
                    words[i] = char.ToUpperInvariant(words[i][0]) + words[i].Substring(1).ToLowerInvariant();
                }
            }
            return string.Join(" ", words);
        }

        private static string ToSnakeCase(string input)
        {
            string[] words = SplitWords(input);
            for (int i = 0; i < words.Length; i++)
            {
                words[i] = words[i].ToLowerInvariant();
            }
            return string.Join("_", words);
        }

        private static string[] SplitWords(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return new string[0];
            }

            // Handle camelCase, PascalCase, snake_case, kebab-case, etc.
            string result = Regex.Replace(input, @"([a-z0-9])([A-Z])", "$1 $2");
            result = Regex.Replace(result, @"[_\-]", " ");
            result = Regex.Replace(result, @"\s+", " ").Trim();

            return result.Split(' ');
        }
    }
}
