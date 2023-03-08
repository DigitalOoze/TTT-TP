using HtmlAgilityPack;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace YellowPagesScrapper
{
    static class Scrapper
    {
        public static List<string> Scrape(string url)
        {

            List<string> result = new List<string>();
            StringBuilder resultString = new StringBuilder("");

            try
            {
                HtmlWeb web = new HtmlWeb();
                HtmlDocument doc = web.Load(url);

                string websiteText = doc.DocumentNode.InnerText;
                MatchCollection emailMatches = Regex.Matches(websiteText, @"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b");

                if (emailMatches.Count == 0)
                {
                    var contactLinks = from e in doc.DocumentNode.Descendants("a")
                                        where e.Attributes["href"] != null && (e.InnerText.ToLower().Contains("contact") || e.InnerText.ToLower().Contains("contact-us") || e.InnerText.ToLower().Contains("kontaktai") || e.InnerText.ToLower().Contains("kontakt"))
                                        select e.Attributes["href"].Value;

                    if (contactLinks.Count() != 0)
                    {
                        // Load the contact page and search for email addresses
                        HtmlDocument contactDoc = web.Load(contactLinks.First());
                        var emailAddresses = from e in contactDoc.DocumentNode.Descendants("a")
                                                where e.Attributes["href"] != null &&
                                                e.Attributes["href"].Value.StartsWith("mailto:")
                                                select e.Attributes["href"].Value;

                        if (emailAddresses.Count() == 0)
                        {
                            // Search for email addresses within the contact page's text
                            string contactText = contactDoc.DocumentNode.InnerText;
                            emailMatches = Regex.Matches(contactText, @"\w+@\w+\.\w+");

                            if (emailMatches.Count != 0)
                            {
                                foreach (Match emailMatch in emailMatches)
                                {
                                    string decodedString = emailMatch.Value;
                                    if (emailMatch.Value.Contains("&#"))
                                        decodedString = HttpUtility.HtmlDecode(emailMatch.Value);

                                    resultString.Append(decodedString + '\n');
                                }
                            }
                        }
                        else
                        {
                            foreach (var email in emailAddresses)
                            {
                                string decodedString = email;
                                if (email.Contains("&#"))
                                    decodedString = HttpUtility.HtmlDecode(email);

                                resultString.Append(decodedString.Substring(7) + '\n');
                            }
                        }
                    }
                }
                else
                {
                    foreach (Match emailMatch in emailMatches)
                    {
                        string decodedString = emailMatch.Value;
                        if (emailMatch.Value.Contains("&#"))
                            decodedString = HttpUtility.HtmlDecode(emailMatch.Value);

                        resultString.Append(decodedString + '\n');
                    }
                }
                result.Add(resultString.ToString());
            }
            catch (HtmlWebException ex)
            {
                Console.WriteLine("Error loading HTML for URL: {0}", url);
                Console.WriteLine(ex.Message);
            }
            catch (UriFormatException ex)
            {
                Console.WriteLine($"An error occurred while processing URL: {url}");
                Console.WriteLine(ex.Message);
                try
                {
                    string newUrl = url + "/index.html";
                    HtmlWeb web = new HtmlWeb();
                    HtmlDocument doc = web.Load(newUrl);
                }
                catch (Exception innerEx)
                {
                    Console.WriteLine($"An error occurred while processing URL: {url}");
                    Console.WriteLine(innerEx.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing URL: {url}");
                Console.WriteLine(ex.Message);
            }
            
            return result;
        }
    }
}