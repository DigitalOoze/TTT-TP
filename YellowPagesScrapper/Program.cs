using HtmlAgilityPack;
using System.Xml.Linq;
using OfficeOpenXml;
using static System.Net.WebRequestMethods;
using YellowPagesScrapper;

Dictionary<string, string> Searching = new Dictionary<string, string>
{
    { "Insurance", "https://www.yellowpages.com/search?search_terms=Insurance&geo_location_terms=CA&page=" },
    { "Investment firm", "https://www.yellowpages.com/search?search_terms=Investment%20firm&geo_location_terms=CA&page=" },
    { "Financial Planning", "https://www.yellowpages.com/search?search_terms=Financial%20Planning%20Consultants&geo_location_terms=CA&page=" },
    { "Brokerage firms", "https://www.yellowpages.com/search?search_terms=brokerage+firm&geo_location_terms=CA&page=" },
    { "Accounting firms", "https://www.yellowpages.com/search?search_terms=accounting+firms&geo_location_terms=CA&page=" },
    { "Venture Capital", "https://www.yellowpages.com/search?search_terms=Venture+Capital&geo_location_terms=CA&page=" },
    { "Real Estate", "https://www.yellowpages.com/search?search_terms=Real+Estate&geo_location_terms=CA&page=" },
    { "car rental", "https://www.yellowpages.com/search?search_terms=car+rental&geo_location_terms=CA&page=" },
    { "lawyer", "https://www.yellowpages.com/search?search_terms=lawyer&geo_location_terms=CA&page=" },
    { "Beauty Salons", "https://www.yellowpages.com/search?search_terms=Beauty+Salons&geo_location_terms=CA&page=" },
    { "legal service plans", "https://www.yellowpages.com/search?search_terms=legal+service+plans&geo_location_terms=CA&page=" },
    { "Skincare", "https://www.yellowpages.com/search?search_terms=skin+care&geo_location_terms=CA&page=" },
    { "Makeup ", "https://www.yellowpages.com/search?search_terms=mAReup&geo_location_terms=CA&page=" },
    { "Spa", "https://www.yellowpages.com/search?search_terms=Spa&geo_location_terms=CA&page=" },
    { "Organic Beauty Skin Care Studio", "https://www.yellowpages.com/search?search_terms=Organic+Beauty+Skin+Care+Studio&geo_location_terms=CA&page=" },
    { "Fitness Coach", "https://www.yellowpages.com/search?search_terms=Fitness+Coach&geo_location_terms=CA&page=" }
};


foreach (var search in Searching)
{
    DateTime StartTime = DateTime.Now;
    int row = 1;
    int NumberOfPages = 0;
    var htmlDoc123 = new HtmlWeb();
    try
    {
        HtmlDocument doc123 = htmlDoc123.Load(search.Value+"1");

        var k = from a in doc123.DocumentNode.Descendants("div")
                 let classAttr = a.Attributes["class"]
                 where classAttr != null && classAttr.Value.Contains("pagination")
                 select a.FirstChild.InnerHtml;

        NumberOfPages = int.Parse(k.First().Substring(16).ToString());
        NumberOfPages = NumberOfPages / 30 + (NumberOfPages % 30 > 0 ? 1 : 0);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Exception: {ex}");
    }
    
    for (int i = 1; i <= NumberOfPages; i++)
    {
        IEnumerable<string> elements = new List<string>();

        Console.WriteLine("Page: " + i + "/" + NumberOfPages + " --- " + search.Key);

        string url = search.Value + i;
        var htmlDoc = new HtmlWeb();
        try
        {
            HtmlDocument doc = htmlDoc.Load(url);
            elements = from a in doc.DocumentNode.Descendants("a")
                       let classAttr = a.Attributes["class"]
                       let hrefAttr = a.Attributes["href"]
                       where classAttr != null && classAttr.Value.Contains("business-name") && hrefAttr != null
                       select hrefAttr.Value;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Page: {i}");
            Console.WriteLine($"Exception: {ex}");
            continue;
        }

        List<Item> items = new List<Item>();

        Parallel.ForEach(elements, element =>
        {
            string href = "https://www.yellowpages.com" + element;

            try
            {
                var htmlDoc1 = new HtmlWeb();
                HtmlDocument doc1 = htmlDoc1.Load(href);

                var companyName = from h in doc1.DocumentNode.Descendants("h1")
                                  let classAttr = h.Attributes["class"]
                                  where classAttr != null && classAttr.Value.Contains("business-name")
                                  select h.InnerText;

                var links = from a in doc1.DocumentNode.Descendants("a")
                            let classAttr = a.Attributes["class"]
                            let hrefAttr = a.Attributes["href"]
                            where classAttr != null && classAttr.Value.Contains("website-link") && hrefAttr != null
                            select hrefAttr.Value;

                var emails = from a in doc1.DocumentNode.Descendants("a")
                             let classAttr = a.Attributes["class"]
                             let hrefAttr = a.Attributes["href"]
                             where classAttr != null && classAttr.Value.Contains("email-business") && hrefAttr != null
                             select hrefAttr.Value;
                
                
                Item item = new Item();
                item.CompanyName = companyName.FirstOrDefault();
                item.Website = links.FirstOrDefault();
                var email = emails.FirstOrDefault();
                if (email != null)
                    item.Email = email.Substring(7);
                else if (item.Website != null)
                    item.Email = Scrapper.Scrape(item.Website).FirstOrDefault();

                items.Add(item);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Page: {i}");
                Console.WriteLine($"Exception: {ex}");
            }
        });

        using (var package = new ExcelPackage(new FileInfo(@"C:\Users\Papuna\Desktop\Result.xlsx")))
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            
            if(i == 1 && package.Workbook.Worksheets[search.Key.ToString()] == null)
                package.Workbook.Worksheets.Add(search.Key.ToString());

            var worksheet = package.Workbook.Worksheets[search.Key];
            if(row == 1)
            {
                worksheet.Cells[row, 1].Value = "CompanyName";
                worksheet.Cells[row, 2].Value = "Website";
                worksheet.Cells[row, 3].Value = "Email";
                row++;
            }

            foreach (var item in items)
            {
                while (worksheet.Cells[row, 1].Value != null || worksheet.Cells[row, 2].Value != null || worksheet.Cells[row, 3].Value != null)
                    row++;

                worksheet.Cells[row, 1].Value = item.CompanyName;
                worksheet.Cells[row, 2].Value = item.Website;
                worksheet.Cells[row, 3].Value = item.Email;

                package.Save();
            }
        }
    }
    Console.WriteLine("\n\n\n\t Link " + search.Key + " time:" + (DateTime.Now - StartTime) + "\n\n\n");
}
