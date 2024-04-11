using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    /// <summary>
    /// Defines a provisioning engine Token. Make sure to only use the TokenContext property to execute queries in token methods.
    /// </summary>
    public abstract class TokenDefinition
    {
        private ClientContext _context;
        protected string CacheValue;
        private readonly string[] _tokens;
        private readonly string[] _unescapedTokens;
        private readonly int _maximumTokenLength;

        /// <summary>
        /// Defines if a token is cacheable and should be added to the token cache during initialization of the token parser. This means that the value for a token will be returned from the cache instead from the GetReplaceValue during the provisioning run. Defaults to true.
        /// </summary>
        public bool IsCacheable { get; set; } = true;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="web">Current site/subsite</param>
        /// <param name="token">token</param>
        public TokenDefinition(Web web, params string[] token)
        {
            Web = web;
            _tokens = token;
            _unescapedTokens = GetUnescapedTokens(token);
            _maximumTokenLength = GetMaximumTokenLength(token);
        }

        /// <summary>
        /// Returns a cloned context which is separate from the current context, not affecting ongoing queries.
        /// </summary>
        public ClientContext TokenContext
        {
            get
            {
                // CHANGED: the URL can be null ...
                if (_context == null && Web.IsPropertyAvailable(w => w.Url))
                {
                    // Make sure that the Url property has been loaded on the web in the constructor
                    _context = Web.Context.Clone(Web.Url);
                }
                return _context;
            }
        }

        /// <summary>
        /// Gets the amount of tokens hold by this token definition
        /// </summary>
        public int TokenCount
        {
            get => _tokens.Length;
        }

        /// <summary>
        /// Gets tokens
        /// </summary>
        /// <returns>Returns array string of tokens</returns>
        public string[] GetTokens()
        {
            return _tokens;
        }

        /// <summary>
        /// Gets the by <see cref="Regex.Unescape"/> processed tokens
        /// </summary>
        /// <returns>Returns array string of by <see cref="Regex.Unescape"/> processed tokens</returns>
        public IReadOnlyList<string> GetUnescapedTokens()
        {
            return _unescapedTokens;
        }

        /// <summary>
        /// Web is a SiteCollection or SubSite
        /// </summary>
        public Web Web { get; set; }

        /// <summary>
        /// Gets array of regular expressions
        /// </summary>
        /// <returns>Returns all Regular Expressions</returns>
        [Obsolete("No longer in use")]
        public Regex[] GetRegex()
        {
            var regexs = new Regex[_tokens.Length];
            for (var q = 0; q < _tokens.Length; q++)
            {
                regexs[q] = new Regex(_tokens[q], RegexOptions.IgnoreCase);
            }
            return regexs;
        }

        /// <summary>
        /// Gets regular expression for the given token
        /// </summary>
        /// <param name="token">token string</param>
        /// <returns>Returns RegularExpression</returns>
        [Obsolete("No longer in use")]
        public Regex GetRegexForToken(string token)
        {
            return new Regex(token, RegexOptions.IgnoreCase);
        }

        /// <summary>
        /// Gets the length of the largest token
        /// </summary>
        /// <returns>Length of the largest token</returns>
        public int GetTokenLength()
        {
            return _maximumTokenLength;
        }

        /// <summary>
        /// abstract method
        /// </summary>
        /// <returns>Returns string</returns>
        public abstract string GetReplaceValue();

        /// <summary>
        /// Clears cache
        /// </summary>
        public void ClearCache()
        {
            CacheValue = null;
        }

        private static int GetMaximumTokenLength(IReadOnlyList<string> tokens)
        {
            var result = 0;

            for (var index = 0; index < tokens.Count; index++)
            {
                result = Math.Max(result, tokens[index].Length);
            }

            return result;
        }

        private static string[] GetUnescapedTokens(IReadOnlyList<string> tokens)
        {
            var result = new string[tokens.Count];

            for (var index = 0; index < tokens.Count; index++)
            {
                result[index] = Regex.Unescape(tokens[index]);
            }

            return result;
        }
    }
}