using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using ClosedXML.Excel;
using HtmlAgilityPack;

namespace ProductParser
{
    internal class Program
    {
        private static readonly HttpClient client = new HttpClient();

        static async Task Main(string[] args)
        {
            // Настройка User-Agent (имитируем браузер)
            client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");

            // Реальный тестовый сайт для парсинга
            string url = "http://books.toscrape.com/";

            // Путь к рабочему столу
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = Path.Combine(desktopPath, "products.xlsx");

            try
            {
                Console.WriteLine($"Загрузка страницы: {url}...");

                // 1. HTTP-запрос
                string htmlContent = await client.GetStringAsync(url);

                // 2. Парсинг HTML
                var htmlDoc = new HtmlDocument();
                htmlDoc.LoadHtml(htmlContent);

                // Ищем заголовки книг
                var productNodes = htmlDoc.DocumentNode.SelectNodes("//h3/a");
                var products = new List<(string Name, string Price)>();

                if (productNodes != null)
                {
                    Console.WriteLine($"Найдено товаров: {productNodes.Count}. Начинаю сбор данных...");

                    foreach (var titleNode in productNodes)
                    {
                        // Название книги
                        string name = titleNode.Attributes["title"]?.Value ?? titleNode.InnerText.Trim();

                        // Цена (ищем в родительском блоке конкретной книги)
                        var priceNode = titleNode.SelectSingleNode("ancestor::div[@class='product_pod']//p[@class='price_color']");
                        string price = priceNode != null ? priceNode.InnerText.Trim() : "Цена не найдена";

                        products.Add((name, price));

                        // Вежливая задержка, чтобы не спамить сервер (полезно для портфолио)
                        await Task.Delay(200);
                    }
                }
                else
                {
                    Console.WriteLine("Товары не найдены. Возможно, изменилась верстка сайта.");
                }

                // 3. Создание Excel файла
                if (products.Count > 0)
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Товары");
                        worksheet.Cell(1, 1).Value = "Название товара";
                        worksheet.Cell(1, 2).Value = "Цена";

                        int row = 2;
                        foreach (var product in products)
                        {
                            worksheet.Cell(row, 1).Value = product.Name;
                            worksheet.Cell(row, 2).Value = product.Price;
                            row++;
                        }

                        worksheet.Columns().AdjustToContents();
                        workbook.SaveAs(filePath);
                    }

                    Console.WriteLine("Парсинг успешно завершен! Файл Excel создан.");
                    Console.WriteLine($"Путь: {filePath}");
                }
                else
                {
                    Console.WriteLine("Нет данных для сохранения.");
                }
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine($"Ошибка сети: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка: {ex.Message}");
            }

            Console.WriteLine("Нажмите любую клавишу для выхода...");
            Console.ReadKey();
        }
    }
}