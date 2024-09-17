using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace WebScrapingTest1
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://www.nbplaza.com.my/260-intel");

            // List to store all products
            List<Product> allProducts = new List<Product>();

            // Loop through each page until there are no more pages
            while (true)
            {
                // Extract products on the current page
                var productElements = driver.FindElements(By.CssSelector(".product-miniature"));
                if (productElements.Count == 0)
                {
                    Console.WriteLine("No products found on this page.");
                    break;
                }

                foreach (var productElement in productElements)
                {
                    // Extract product details
                    string? productName = productElement.FindElement(By.CssSelector(".product_name")).Text;
                    string? regularPrice = null;
                    try
                    {
                        regularPrice = productElement.FindElement(By.CssSelector(".regular-price")).Text;
                    }
                    catch { }
                    string? salePrice = null;
                    try
                    {
                        salePrice = productElement.FindElement(By.CssSelector(".price-sale")).Text;
                    }
                    catch
                    {
                        try
                        {
                            salePrice = productElement.FindElement(By.CssSelector(".price")).Text;
                        }
                        catch { }
                    }
                    string? discount = null;
                    try
                    {
                        discount = productElement.FindElement(By.CssSelector(".discount-amount")).Text;

                    }
                    catch { }
                    var imageUrl = productElement.FindElement(By.CssSelector(".first-image")).GetAttribute("data-src");

                    // Add the product to the list
                    allProducts.Add(new Product
                    {
                        Name = productName,
                        RegularPrice = regularPrice,
                        SalePrice = salePrice,
                        Discount = discount,
                        ImageUrl = imageUrl
                    });
                }

                // Try to find the "Next" button and navigate to the next page
                try
                {
                    var nextButton = driver.FindElement(By.CssSelector(".pagination .next a")); // Update the selector as needed
                    if (nextButton.Displayed)
                    {
                        nextButton.Click();
                        Thread.Sleep(3000); // Wait for the page to load
                    }
                    else
                    {
                        Console.WriteLine("No more pages to scrape.");
                        break;
                    }
                }
                catch (NoSuchElementException)
                {
                    Console.WriteLine("No 'Next' button found. Ending pagination.");
                    break;
                }

            }

            // Output the scraped information
            foreach (var product in allProducts)
            {
                Console.WriteLine("======================================");
                Console.WriteLine($"Product Name: {product.Name}");
                Console.WriteLine($"Regular Price: {product.RegularPrice}");
                Console.WriteLine($"Sale Price: {product.SalePrice}");
                Console.WriteLine($"Discount: {product.Discount}");
                Console.WriteLine($"Image URL: {product.ImageUrl}");
                Console.WriteLine("======================================");
            }

            // Export data to Excel using ClosedXML
            ExportToExcel(allProducts);

            // Close the browser
            driver.Quit();

            Console.WriteLine();
        }

        // Method to export data to Excel
        private static void ExportToExcel(List<Product> products)
        {
            string filePath = Path.Combine(Environment.CurrentDirectory, "Products.xlsx");

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Products");

                // Adding header row
                worksheet.Cell(1, 1).Value = "Product Name";
                worksheet.Cell(1, 2).Value = "Regular Price";
                worksheet.Cell(1, 3).Value = "Sale Price";
                worksheet.Cell(1, 4).Value = "Discount";
                worksheet.Cell(1, 5).Value = "Image URL";

                // Adding product details
                for (int i = 0; i < products.Count; i++)
                {
                    worksheet.Cell(i + 2, 1).Value = products[i].Name;
                    worksheet.Cell(i + 2, 2).Value = products[i].RegularPrice;
                    worksheet.Cell(i + 2, 3).Value = products[i].SalePrice;
                    worksheet.Cell(i + 2, 4).Value = products[i].Discount;
                    worksheet.Cell(i + 2, 5).Value = products[i].ImageUrl;
                }

                // Save the Excel file
                workbook.SaveAs(filePath);
                Console.WriteLine($"Excel file saved at: {filePath}");
            }
        }
    }

    // Helper class to store product info
    class Product
    {
        public string Name { get; set; }
        public string RegularPrice { get; set; }
        public string SalePrice { get; set; }
        public string Discount { get; set; }
        public string ImageUrl { get; set; }
    }
}
