using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Text;

Console.OutputEncoding = Encoding.UTF8;

string keyword = "Tân Phú";
Console.WriteLine($"🔍 Tìm nhóm với từ khóa: {keyword}");

string userDataDir = @"C:\SeleniumProfiles\MyProfile";
string profileDir = "Default";

var options = new ChromeOptions();
options.AddArgument($"--user-data-dir={userDataDir}");
options.AddArgument($"--profile-directory={profileDir}");
options.AddArgument("--start-maximized");
options.AddArgument("--disable-blink-features=AutomationControlled");
options.AddExcludedArgument("enable-automation");
options.AddAdditionalOption("useAutomationExtension", false);

using var driver = new ChromeDriver(options);

string searchUrl = $"https://www.facebook.com/search/groups?q={Uri.EscapeDataString(keyword)}";
driver.Navigate().GoToUrl(searchUrl);
Thread.Sleep(5000);

// Cuộn trang để tải nhiều nhóm hơn
for (int i = 0; i < 10; i++)
{
    ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(0, document.body.scrollHeight);");
    Thread.Sleep(1500);
}

// 👉 Bước 1: Chỉ lấy text + href, không giữ WebElement
var rawLinks = driver.FindElements(By.XPath("//a[contains(@href, '/groups/')]"));
var results = new List<(string name, string url)>();

foreach (var link in rawLinks)
{
    string href = link.GetAttribute("href");
    string name = link.Text.Trim();

    if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(href))
    {
        if (!results.Any(x => x.url == href))
        {
            results.Add((name, href));
            Console.WriteLine($"📌 {name} - {href}");
        }
    }
}

// 👉 Bước 2: Duyệt từng nhóm và bấm nút "Tham gia"
foreach (var (name, url) in results)
{
    try
    {
        Console.WriteLine($"\n🔗 Mở nhóm: {name}");
        driver.Navigate().GoToUrl(url);
        Thread.Sleep(5000);

        try
        {
            var joinBtn = driver.FindElement(By.XPath("//span[text()='Tham gia nhóm']"));
            if (joinBtn != null)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", joinBtn);
                Thread.Sleep(1000);
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", joinBtn);
                Console.WriteLine("✅ Đã bấm nút 'Tham gia'");
                Thread.Sleep(3000);
            }
        }
        catch (NoSuchElementException)
        {
            Console.WriteLine("⚠️ Không thấy nút 'Tham gia'");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"❌ Lỗi khi xử lý nhóm {name}: {ex.Message}");
    }
}

// Xuất Excel
var workbook = new XLWorkbook();
var worksheet = workbook.Worksheets.Add("Groups");
worksheet.Cell(1, 1).Value = "Tên nhóm";
worksheet.Cell(1, 2).Value = "Link";

int row = 2;
foreach (var (name, url) in results)
{
    worksheet.Cell(row, 1).Value = name;
    worksheet.Cell(row, 2).Value = url;
    row++;
}

string filePath = $"FacebookGroups_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
workbook.SaveAs(filePath);

Console.WriteLine($"\n📁 Đã lưu kết quả vào: {filePath}");
driver.Quit();
Console.WriteLine("\n✅ Hoàn tất.");
