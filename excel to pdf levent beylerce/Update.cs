using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

class Update
{
    static string CurrentVersion = "1.0.0"; // bu senin şu anki versiyonun

    static async Task Main()
    {
        string latest = await GetLatestVersion();
        Console.WriteLine("Latest version: " + latest);

        if (latest != CurrentVersion)
        {
            Console.WriteLine("Yeni sürüm var: " + latest);
            string zipPath = await DownloadLatestAsset();
            Console.WriteLine("Güncelleme indirildi: " + zipPath);
            // Burada updater / unzip / çalıştır vs işlemlerini yapabilirsin
        }
        else
        {
            Console.WriteLine("Zaten en son sürümdesin.");
        }
    }

    static async Task<string> GetLatestVersion()
    {
        using (HttpClient client = new HttpClient())
        {
            client.DefaultRequestHeaders.Add("User-Agent", "MyAppUpdater");
            string url = "https://api.github.com/repos/Albayrak050/excel-to-pdf-levent-beylerce/releases/latest";
            var json = await client.GetStringAsync(url);
            var obj = JObject.Parse(json);
            return (string)obj["tag_name"];
        }
    }

    static async Task<string> DownloadLatestAsset()
    {
        using (HttpClient client = new HttpClient())
        {
            client.DefaultRequestHeaders.Add("User-Agent", "MyAppUpdater");
            string url = "https://api.github.com/repos/Albayrak050/excel-to-pdf-levent-beylerce/releases/latest";
            var json = await client.GetStringAsync(url);
            var obj = JObject.Parse(json);
            var asset = obj["assets"]?[0];
            if (asset == null)
                throw new Exception("Asset bulunamadı.");

            string downloadUrl = (string)asset["browser_download_url"];
            string localPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excel.to.pdf.levent.beylerce.exe");
            var bytes = await client.GetByteArrayAsync(downloadUrl);
            File.WriteAllBytes(localPath, bytes);
            return localPath;
        }
    }
}
