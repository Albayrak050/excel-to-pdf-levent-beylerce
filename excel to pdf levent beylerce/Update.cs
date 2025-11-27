using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

public class Updater
{
    private readonly string CurrentVersion = "1.0.0"; // programının versiyonu

    public async Task CheckUpdate()
    {
        string latest = await GetLatestVersion();

        if (latest != CurrentVersion)
        {
            System.Windows.Forms.MessageBox.Show(
                $"Yeni sürüm bulundu: {latest}\nGüncelleme indiriliyor..."
            );

            string zipPath = await DownloadLatestAsset();

            System.Windows.Forms.MessageBox.Show(
                $"Güncelleme indirildi:\n{zipPath}\nZip'i açıp kurulumu yapmak size kaldı."
            );
        }
        else
        {
            System.Windows.Forms.MessageBox.Show("Zaten en son sürüm kullanılıyor.");
        }
    }

    private async Task<string> GetLatestVersion()
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

    private async Task<string> DownloadLatestAsset()
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

            string localPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excel to pdf levent beylerce.exe");

            var bytes = await client.GetByteArrayAsync(downloadUrl);
            File.WriteAllBytes(localPath, bytes);

            return localPath;
        }
    }
}
