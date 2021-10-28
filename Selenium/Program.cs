using Microsoft.Win32;
using RestSharp;
using System.Diagnostics;
using System.IO.Compression;
using System.IO.Packaging;
using System.Text.RegularExpressions;
using System.Web;

var BaseDirectory = AppDomain.CurrentDomain.BaseDirectory;
var chromedriverPath = Path.Combine(BaseDirectory, "chromedriver.exe");
var url = "http://npm.taobao.org/mirrors/chromedriver/";


if(!check_update_chromedriver())
{
    Console.WriteLine("当前电脑环境存在问题，请解决后在执行");
    Console.ReadLine();
    return;
}



using (var driver = new OpenQA.Selenium.Chrome.ChromeDriver())
{
    driver.Navigate().GoToUrl("http://www.baidu.com");
    var source = driver.PageSource;
    Console.WriteLine(source);
    Console.ReadLine();
}

/// <summary>
/// 检查驱动更新
/// </summary>
bool check_update_chromedriver()
{
    var chromeVersion = get_Chrome_version();
    if (string.IsNullOrEmpty(chromeVersion))
    {
        Console.WriteLine("未安装Chrome，请在GooGle Chrome官网：https://www.google.cn/chrome/ 下载。");
        return false;
    }
    Console.WriteLine("当前Chrome版本为：" + chromeVersion);

    var chrome_main_version = int.Parse(chromeVersion.Split(".")[0]); // chrome主版本号
    var driverVersion = get_version();
    var driver_main_version = 0;
    if (string.IsNullOrEmpty(driverVersion))
    {
        Console.WriteLine("未安装Chromedriver，正在为您自动下载>>>");
        if (!download_lase_driver(chromeVersion, chrome_main_version))
        {
            return false;
        }
    }
    driver_main_version = int.Parse(driverVersion.Split(".")[0]);
    if (driver_main_version != chrome_main_version)
    {
        Console.WriteLine("chromedriver版本与chrome浏览器不兼容，更新中>>>");
        if (!download_lase_driver(chromeVersion, chrome_main_version))
        {
            return false;
        }
    }

    Console.WriteLine("chromedriver版本已与chrome浏览器相兼容，无需更新chromedriver版本！");
    return true;
}

/// <summary>
/// 下载最新驱动
/// </summary>
bool download_lase_driver(string chromeVersion, int chrome_main_version)
{
    var download_url = "";
    var versionList = get_server_chrome_versions();
    if (versionList.Any(m => { return m == chromeVersion; }))
    {
        download_url = $"{url}{chromeVersion}/chromedriver_win32.zip";
    }
    else
    {
        foreach (var version in versionList)
        {
            if (version.StartsWith(chrome_main_version.ToString()))
            {
                download_url = $"{url}{version}/chromedriver_win32.zip";
                break;
            }
        }
        if (string.IsNullOrEmpty(download_url))
        {
            Console.WriteLine("暂无法找到与chrome兼容的chromedriver版本，请在http://npm.taobao.org/mirrors/chromedriver/ 核实。");
            return false;
        }
    }
    download_driver(download_url);
    unzip_driver(BaseDirectory);
    var driverVersion = get_version();
    if (string.IsNullOrEmpty(driverVersion))
    {
        Console.WriteLine("驱动安装失败,请尝试更新最新Chrome浏览器再试");
        return false;
    }
    return true;
}

/// <summary>
/// 下载文件
/// </summary>
void download_driver(string download_url)
{
    var client = new RestClient(download_url);
    var request = new RestRequest(Method.GET);
    var file = client.DownloadData(request);
    File.WriteAllBytes("chromedriver.zip", file);
    Console.WriteLine("下载成功");
}

/// <summary>
/// 解压驱动
/// </summary>
void unzip_driver(string path)
{
    ZipFile.ExtractToDirectory("chromedriver.zip", path);
    File.Delete("chromedriver.zip");
}

/// <summary>
/// 获取网络谷歌浏览器驱动 版本列表
/// </summary>
List<string> get_server_chrome_versions()
{
    var versionList = new List<string>();
    var url = "http://npm.taobao.org/mirrors/chromedriver/";
    var client = new RestClient(url);
    var rep = client.Get<string>(new RestRequest(Method.GET));
    string pattern = @"\d\/"">(.*?)\/<\/a>.*?Z";
    string input = rep.Content;
    RegexOptions options = RegexOptions.Multiline;

    foreach (Match m in Regex.Matches(input, pattern, options))
    {
        versionList.Add(m.Groups[1].Value);
    }
    return versionList;
}

/// <summary>
/// 读取注册表 获取本机谷歌浏览器版本
/// </summary>
string? get_Chrome_version()
{
    var rk = Registry.CurrentUser;
    if (rk == null) return null;
    var key = rk.OpenSubKey(@"Software\Google\Chrome\BLBeacon");
    if (key == null)
        return null;
    var version = key.GetValue("version");
    if (version == null)
        return null;
    return version.ToString();
}

/// <summary>
/// 获取chromedriver版本
/// </summary>
string get_version()
{
    try
    {
        var outStr = ExecuteCmd("chromedriver --version");
        outStr = outStr.Split(' ')[1];
        return outStr;
    }
    catch
    {
        return "";
    }
}

/// <summary>
/// 执行cmd 获取返回值
/// </summary>
string ExecuteCmd(string cmd)
{
    Process process = new Process();
    process.StartInfo.FileName = "cmd.exe";
    process.StartInfo.UseShellExecute = false;   // 是否使用外壳程序 
    process.StartInfo.CreateNoWindow = true;   //是否在新窗口中启动该进程的值 
    process.StartInfo.RedirectStandardInput = true;  // 重定向输入流 
    process.StartInfo.RedirectStandardOutput = true;  //重定向输出流 
    process.StartInfo.RedirectStandardError = true;  //重定向错误流 
    var strCmd = cmd + " &exit";
    process.Start();
    process.StandardInput.WriteLine(strCmd);
    var stringoutput = process.StandardOutput.ReadToEnd();//获取输出信息 
    stringoutput = stringoutput.Substring(stringoutput.IndexOf(strCmd) + strCmd.Length);
    process.WaitForExit();
    int n = process.ExitCode;  // n 为进程执行返回值 
    process.Close();
    return stringoutput.Trim();
}
