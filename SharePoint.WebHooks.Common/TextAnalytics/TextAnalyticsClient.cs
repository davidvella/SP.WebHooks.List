using Microsoft.Azure;
using Microsoft.Azure.CognitiveServices.Language.TextAnalytics;
using Microsoft.Azure.CognitiveServices.Language.TextAnalytics.Models;
using Microsoft.Rest;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace SharePoint.WebHooks.Common.TextAnalytics
{
    /// <summary>
    ///  The Text Analytics API is a suite of text analytics web services built with best-in-class
    ///  Microsoft machine learning algorithms. The API can be used to analyze unstructured
    ///  text for tasks such as sentiment analysis, key phrase extraction and language
    ///  detection. No training data is needed to use this API; just bring your text data.
    ///  This API uses advanced natural language processing techniques to deliver best
    ///  in class predictions. Further documentation can be found in https://docs.microsoft.com/en-us/azure/cognitive-services/text-analytics/overview
    /// </summary>
    class TextAnalyticsClient : Microsoft.Azure.CognitiveServices.Language.TextAnalytics.TextAnalyticsClient
    {
        public TextAnalyticsClient() : base(new ApiKeyServiceClientCredentials(CloudConfigurationManager.GetSetting("TextAnalyticsApiKey")))
        {
            Endpoint = CloudConfigurationManager.GetSetting("TextAnalyticsEndpoint");
        }

        /// <summary>
        /// The API returns a list of strings denoting the key talking points in the input text.
        /// </summary>
        /// <param name="text"> </param>
        /// <param name="showStats">(optional) if set to true, response will contain input and document level statistics.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <remarks>See the <a href="https://docs.microsoft.com/en-us/azure/cognitive-services/text-analytics/overview#supported-languages">Text Analytics Documentation</a> for details about the languages that are supported by key phrase extraction.</remarks>
        public async Task<KeyPhraseBatchResult> KeyPhrasesStringAsync(string text, bool? showStats = null, CancellationToken cancellationToken = default)
        {
            // Get Sentences from string
            var sentences = FormatTextToSentences(text);

            // Get language of inputs
            var languageInput = new List<LanguageInput>();
            var i = 0;
            foreach (var s in sentences)
            {
                if (s.Length > 10) languageInput.Add(new LanguageInput(id: i.ToString(), text: s));
                i++;
            }
            var langResults = await this.DetectLanguageAsync(showStats, new LanguageBatchInput(languageInput), cancellationToken);

            // Make Key Phrases Service
            var inputDocuments = new List<MultiLanguageInput>();
            foreach (var doc in langResults.Documents)
            {
                inputDocuments.Add(new MultiLanguageInput(doc.DetectedLanguages.First().Iso6391Name, doc.Id, languageInput.First(inp => inp.Id == doc.Id).Text));
            }

            return await this.KeyPhrasesAsync(null, new MultiLanguageBatchInput(inputDocuments), cancellationToken);
        }

        private static List<string> FormatTextToSentences(string text)
        {
            // sanitize text a bit
            text = Regex.Replace(text, @"[\r\n\t\f\v]", " ");
            // remove extremely long words - they'll be headers, malformed parts or urls
            text = Regex.Replace(text, @"\S{30,}", " ", RegexOptions.None);
            // remove numbers and everything else but text.
            text = Regex.Replace(text, @"[^a-zA-Z.,'!?äöåü]", " ", RegexOptions.IgnoreCase);
            // lastly, remove extra whitespace
            text = Regex.Replace(text, @"( +)", " ");

            var regExSentenceDelimiter = new Regex(@"(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?)\s");
            var sentences = regExSentenceDelimiter.Split(text).ToList();

            // figure out, which sentence length we're using based on set accuracylevel. The default value is 5120 (set by the API)
            const int limit = 2560;

            var finalizedSentences = new List<string>();

            var sentenceCandidate = "";
            foreach (var sentence in sentences)
            {
                // SANITIZE AND SPLIT
                // drop short sentences (they'll be like "et al", one-liners like "go figure" or just "."
                if (sentence.Length < 10) continue;

                // combine or add other sentences
                if (sentenceCandidate.Length + sentence.Length > limit)
                {
                    finalizedSentences.Add(sentenceCandidate);
                    sentenceCandidate = sentence;
                }
                else
                {
                    sentenceCandidate += " " + sentence;
                }
            }
            // finally, add the last candidate
            finalizedSentences.Add(sentenceCandidate);

            return finalizedSentences;
        }

    }
        
}
