using Serilog;
using System.Text.RegularExpressions;

namespace MassBlattLinkerr.Core
{
    /// <summary>
    /// Class for handling operations related to articles and their versions.
    /// </summary>
    public class ArticleHandler
    {
        private readonly ILogger? _logger;

        public ArticleHandler(ILogger? logger = null)
        {
            if (logger is not null)
            {
                _logger = logger;
            }
        }
        /// <summary>
        /// Extracts article name and version from the given input string.
        /// </summary>
        /// <param name="input">Input string containing article and version.</param>
        /// <returns>Tuple containing article and version.</returns>
        public (string, string) ExtractArticleAndVersion(string input)
        {
            var match = Regex.Match(input, @"(\d+)(?:\((\w*)\))?\s*(\d*)");

            if (match.Success)
            {
                var article = match.Groups[1].Value;
                string version = string.IsNullOrEmpty(match.Groups[3].Value) ? match.Groups[2].Value : match.Groups[3].Value;

                return (article, version);
            }

            _logger?.Error("Nepodařilo se extrahovat artikl a tiskovou verzi.");
            return ("", "");
        }

        /// <summary>
        /// Extracts article name, print version, and file path from the given file path.
        /// </summary>
        /// <param name="filePath">File path to extract data from.</param>
        /// <returns>Tuple containing article name, print version, and file path if successful; otherwise, null.</returns>
        public (string ArticleName, string PrintVersionName, string FilePath)? ExtractDataFromPath(string filePath)
        {
            string articlePattern = @"\\(\d{4,5})\\[^\\]*$";
            string versionPattern = @"\.\d{2}(\d{3})[^\\]*\.pdf$";

            var articleMatch = Regex.Match(filePath, articlePattern);
            var articleName = articleMatch.Success ? articleMatch.Groups[1].Value : string.Empty;

            var versionMatch = Regex.Match(filePath, versionPattern);
            var printVersion = versionMatch.Success ? versionMatch.Groups[1].Value : string.Empty;

            if (!string.IsNullOrEmpty(articleName) && !string.IsNullOrEmpty(printVersion))
            {
                return (articleName, printVersion, filePath);
            }

            _logger?.Warning("Nepodařilo se extrahovat data z cesty. {path}", filePath);
            return null;
        }
    }
}
