using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace NewBingBackground
{
    class Program
    {

        private static void Main(string[] args)
        {
            var urlBase = GetBackgroundUrlBase();
            Image background = DownloadBackground(urlBase + GetResolutionExtension(urlBase));
            SaveBackground(background);
            SetDesktopBackground(background, ImagePosition.Fill);
        }


        private static dynamic DownloadJSON()
        {
            using (WebClient webClient = new WebClient())
            {
                Console.WriteLine("Downloading JSON...");
                var jsonString = webClient.DownloadString("https://www.bing.com/HPImagearchive.aspx?format=js&idx=0&n=1&mkt=pl-PL");
                return JsonConvert.DeserializeObject<dynamic>(jsonString);
            }
        }

        private static string GetBackgroundUrlBase()
        {
            var jsonObject = DownloadJSON();
            return "https://www.bing.com" + jsonObject.images[0].urlbase;
        }

        private static string GetTitleOfBackground()
        {
            var jsonObject = DownloadJSON();
            var copyrightText = jsonObject.images[0].copyright;
            return copyrightText.Substring(0, copyrightText.IndexOf(" ("));
        }

        private static bool ExistingWebsiteUrl(string url)
        {
            try
            {
                WebRequest request = WebRequest.Create(url);
                request.Method = "MAIN";
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                return response.StatusCode == HttpStatusCode.OK;
            }
            catch
            {
                return false;
            }
        }

        private static string GetResolutionExtension(string url)
        {
            Rectangle resolution = Screen.PrimaryScreen.Bounds;
            var widthXHeight = resolution.Width + "x" + resolution.Height;
            var potentialExtension = "_" + widthXHeight + ".jpg";
            if (ExistingWebsiteUrl(url + potentialExtension))
            {
                Console.WriteLine("Background for {0} found!", widthXHeight);
                return potentialExtension;
            }
            else
            {
                Console.WriteLine("No background for {0} was found! \nUsing 1920x1080 instead.", widthXHeight);
                return "_1920x1080.jpg";
            }
        }

        private static Image DownloadBackground(string url)
        {
            Console.WriteLine("Downloading background...");
            WebRequest request = WebRequest.Create(url);
            WebResponse reponse = request.GetResponse();
            Stream stream = reponse.GetResponseStream();
            return Image.FromStream(stream);
        }

        private static string GetBackgroundImagePath()
        {
            var directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "Bing Backgrounds", DateTime.Now.Year.ToString());
            Directory.CreateDirectory(directory);
            return Path.Combine(directory, DateTime.Now.ToString("M-d-yyyy") + ".bmp");
        }

        private static void SaveBackground(Image background)
        {
            Console.WriteLine("Saving background...");
            background.Save(GetBackgroundImagePath(), System.Drawing.Imaging.ImageFormat.Bmp);
        }

        private enum ImagePosition
        {
            Tile,
            Center,
            Stretch,
            Fit,
            Fill
        }

        internal sealed class NativeMethods
        {
            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            internal static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
        }

        private static void SetDesktopBackground(Image background, ImagePosition style)
        {
            Console.WriteLine("Setting desktop background...");
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(Path.Combine("Control Panel", "Desktop"), true))
            {
                switch (style)
                {
                    case ImagePosition.Tile:
                        key.SetValue("ImagePosition", "0");
                        key.SetValue("TileWallpaper", "1");
                        break;
                    case ImagePosition.Center:
                        key.SetValue("ImagePosition", "0");
                        key.SetValue("TileWallpaper", "0");
                        break;
                    case ImagePosition.Stretch:
                        key.SetValue("ImagePosition", "2");
                        key.SetValue("TileWallpaper", "0");
                        break;
                    case ImagePosition.Fit:
                        key.SetValue("ImagePosition", "6");
                        key.SetValue("TileWallpaper", "0");
                        break;
                    case ImagePosition.Fill:
                        key.SetValue("ImagePosition", "10");
                        key.SetValue("TileWallpaper", "0");
                        break;
                }
            }
            const int SetBackground = 20;
            const int UpdateIniFile = 1;
            const int SendWindowsIniChange = 2;
            NativeMethods.SystemParametersInfo(SetBackground, 0, GetBackgroundImagePath(), UpdateIniFile | SendWindowsIniChange);
        }
    }
}
