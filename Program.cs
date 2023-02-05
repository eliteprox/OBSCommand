using System;
using System.IO;
using System.Text;
using System.Globalization;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Threading.Tasks;
using Microsoft.VisualBasic;

using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Chrome;

using OBSWebsocketDotNet;
using Tweetinvi;
using System.Web;


namespace OBSCommand {
    static class Program {
        private static OBSWebsocket _obs = new OBSWebsocket();
        private static bool isConnected = false;
        private static string disconnectInfo = "";
        private static string[] twittertokens;

        public static bool tweet { get; private set; }
        public static double runtime { get; private set; }
        public static string tweetmsg { get; private set; }
        public static string channel_url { get; private set; }
        public static bool showdate { get; private set; }
        public static string hashtag { get; private set; }

        

        public static async Task Main(string[] args) {
            args = Environment.GetCommandLineArgs();
            string password = "";
            string server = "ws://127.0.0.1:4455";

            string profile = "";
            string scene = "";
            string hidesource = "";
            string showsource = "";
            string togglesource = "";
            string toggleaudio = "";
            string mute = "";
            string unmute = "";
            string fadeopacity = "";
            string slidesetting = "";
            string slideasync = "";
            string setvolume = "";
            bool stopstream = false;
            bool startstream = false;
            bool startrecording = false;
            bool stoprecording = false;
            string sendjson = "";
            string command = "";
            double setdelay = 0;
            string description = "";
            string title = "";

        TextWriter myout = Console.Out;
            StringBuilder builder = new StringBuilder();
            TextWriter writer = new StringWriter(builder);
            Console.SetOut(writer);

            string errormessage = "";

            if (args.Count() == 1) {
                PrintUsage();
                System.Environment.Exit(0);
            }

            bool isInitialized = false;
            bool skipfirst = false;
            int argsindex = 0;

            _obs.Connected += Connect;
            _obs.Disconnected += Disconnect;

            foreach (string arg in args) {
                argsindex += 1;
                if (skipfirst == false) {
                    skipfirst = true;
                    continue;
                }

                if (arg == "?" | arg == "/?" | arg == "-?" | arg == "help" | arg == "/help" | arg == "-help") {
                    PrintUsage();
                    System.Environment.Exit(0);
                }

                if (arg.StartsWith("/server=")) {
                    server = "ws://" + arg.Replace("/server=", "");
                    continue; // get credentials first before trying to connect!
                }
                if (arg.StartsWith("/password=")) {
                    password = arg.Replace("/password=", "");
                    continue; // get credentials first before trying to connect!
                }

                if (arg.StartsWith("/setdelay=")) {
                    string tmp = arg.Replace("/setdelay=", "");
                    tmp = tmp.Replace(",", ".");
                    if (!Information.IsNumeric(tmp)) {
                        Console.WriteLine("Error: setdelay is not numeric");
                        continue;
                    } else {
                        setdelay = double.Parse(tmp, System.Globalization.CultureInfo.InvariantCulture);
                        continue;
                    }
                }

                if (arg.StartsWith("/delay=")) {
                    string tmp = arg.Replace("/delay=", "");
                    tmp = tmp.Replace(",", ".");
                    if (!Information.IsNumeric(tmp)) {
                        Console.WriteLine("Error: delay is not numeric");
                        continue;
                    } else {
                        double timeDelay; double.TryParse(tmp, out timeDelay);
                        Thread.Sleep((int)Math.Round(timeDelay) * 1000);
                        continue;
                    }
                }

                if (arg.StartsWith("/profile="))
                    profile = arg.Replace("/profile=", "");
                if (arg.StartsWith("/scene="))
                    scene = arg.Replace("/scene=", "");
                if (arg.StartsWith("/command="))
                    command = arg.Replace("/command=", "");
                if (arg.StartsWith("/sendjson="))
                    sendjson = arg.Replace("/sendjson=", "");
                if (arg.StartsWith("/hidesource="))
                    hidesource = arg.Replace("/hidesource=", "");
                if (arg.StartsWith("/showsource="))
                    showsource = arg.Replace("/showsource=", "");
                if (arg.StartsWith("/togglesource="))
                    togglesource = arg.Replace("/togglesource=", "");
                if (arg.StartsWith("/toggleaudio="))
                    toggleaudio = arg.Replace("/toggleaudio=", "");
                if (arg.StartsWith("/mute="))
                    mute = arg.Replace("/mute=", "");
                if (arg.StartsWith("/unmute="))
                    unmute = arg.Replace("/unmute=", "");
                if (arg.StartsWith("/setvolume="))
                    setvolume = arg.Replace("/setvolume=", "");
                if (arg.StartsWith("/fadeopacity="))
                    fadeopacity = arg.Replace("/fadeopacity=", "");
                if (arg.StartsWith("/slidesetting="))
                    slidesetting = arg.Replace("/slidesetting=", "");
                if (arg.StartsWith("/slideasync="))
                    slideasync = arg.Replace("/slideasync=", "");
                if (arg == "/startstream")
                    startstream = true;
                if (arg == "/stopstream")
                    stopstream = true;
                if (arg == "/startrecording")
                    startrecording = true;
                if (arg == "/stoprecording")
                    stoprecording = true;


                if (arg.StartsWith("/runtime=")) {
                    string strRuntime = arg.Replace("/runtime=", "");
                    if (strRuntime.Contains(".")) {
                        string[] arrStopTime = strRuntime.Split('.');
                        int minutes; int.TryParse(arrStopTime[0], out minutes);
                        int seconds; int.TryParse(arrStopTime[1], out seconds);
                        runtime = new TimeSpan(0, 0, minutes, seconds, 0).TotalSeconds;
                    } else {
                        int minutes; int.TryParse(strRuntime, out minutes);
                        runtime = new TimeSpan(0, 0, minutes, 0, 0).TotalSeconds;
                    }
                }

                if (arg.StartsWith("/tweet=")) {
                    //Multi-token
                    twittertokens = arg.Replace("/tweet=", "").Split('|');
                    if (twittertokens.Count() > 0) {
                        tweet = true;
                    }
                }

                if (arg.StartsWith("/msg=")) {
                    tweetmsg = arg.Replace("/msg=", "");
                }

                if (arg.StartsWith("/channelurl=")) {
                    channel_url = arg.Replace("/channelurl=", "");
                }

                if (arg.StartsWith("/description=")) {
                    description = arg.Replace("/description=", "");
                }

                if (arg.StartsWith("/title=")) {
                    title = arg.Replace("/title=", "");
                }

                if (arg.StartsWith("/showdate")) {
                    showdate = true;
                }
                if (arg.StartsWith("/hashtag=")) {
                    hashtag = arg.Replace("/hashtag=", "");
                }
            }

            try {
                if (isInitialized == false) {
                    isInitialized = true;
                    _obs.WSTimeout = new TimeSpan(0, 0, 0, 3);
                    _obs.ConnectAsync(server, password);
                    int i = 0;
                    while (!isConnected) {
                        Thread waitThread = new Thread(() =>
                        {
                            Thread.Sleep(10);
                        });
                        waitThread.Start();
                        waitThread.Join();
                        i += 1;
                        if (i > 300) {
                            Console.Write("Error: can't connect to OBS websocket plugin!");
                            System.Environment.Exit(0);
                        }
                        if (disconnectInfo != "") {
                            Console.Write("Error: " + disconnectInfo);
                            System.Environment.Exit(0);
                        }
                    }
                }

                if (profile != "") {
                    JObject fields = new JObject();
                    fields.Add("profileName", profile);
                    _obs.SendRequest("SetCurrentProfile", fields);
                }

                if (scene != "") {
                    JObject fields = new JObject();
                    fields.Add("sceneName", scene);
                    JObject response = _obs.SendRequest("SetCurrentProgramScene", fields);
                }

                // sendjson
                if (sendjson != "") {
                    JObject json = new JObject();
                    if (!sendjson.Contains("=")) {
                        errormessage = "sendjson missing \"=\" after command";
                    }
                    string[] tmp = new[] { "", "" };
                    try {
                        tmp[0] = sendjson.Substring(0, sendjson.IndexOf("="));
                        tmp[1] = sendjson.Substring(sendjson.IndexOf("=") + 1);
                        tmp[1] = tmp[1].Replace(Strings.Chr(39), Strings.Chr(34));
                        json = JObject.Parse(tmp[1]);
                        Console.WriteLine(_obs.SendRequest(tmp[0], json));
                    } catch (Exception ex) {
                        errormessage = "sendjson error:" + Constants.vbCrLf + ex.Message.ToString();
                    }
                }

                if (command != "") {
                    command = command.Replace("'", "\"");

                    object myParameter = null;

                    try {
                        if (command.Contains(",")) {
                            string[] tmp = command.Split(",");

                            JObject fields = new JObject();
                            for (var a = 1; a <= tmp.Count() - 1; a++) {
                                string[] tmpsplit = SplitWhilePreservingQuotedValues(tmp[a], '=', true);
                                if (tmpsplit.Count() < 2)
                                    Console.WriteLine("Error with command \"" + command + "\": " + "Missing a = in Name=Type");

                                if (tmpsplit.Count() > 2) {
                                    JObject subfield = new JObject();
                                    subfield.Add(ConvertToType(tmpsplit[1]).ToString(), ConvertToType(tmpsplit[2]));
                                    fields.Add(ConvertToType(tmpsplit[0]).ToString(), subfield);
                                } else
                                    fields.Add(ConvertToType(tmpsplit[0]).ToString(), ConvertToType(tmpsplit[1]));
                            }

                            Console.WriteLine(_obs.SendRequest(tmp[0], fields));
                        } else
                            Console.WriteLine(_obs.SendRequest(command));
                    } catch (Exception ex) {
                        Console.WriteLine("Error with command \"" + command + "\": " + ex.Message.ToString());
                    }
                }
                if (hidesource != "") {
                    if (hidesource.Contains("/")) {
                        string[] tmp = hidesource.Split("/");

                        // scene/source
                        if (tmp.Count() == 2) {
                            JObject fields = new JObject();
                            fields.Add("sceneName", tmp[0]);
                            fields.Add("sceneItemId", GetSceneItemId(tmp[0], tmp[1]));
                            fields.Add("sceneItemEnabled", false);
                            _obs.SendRequest("SetSceneItemEnabled", fields);
                        }
                    } else {
                        var CurrentScene = GetCurrentProgramScene();
                        JObject fields = new JObject();
                        fields.Add("sceneName", CurrentScene);
                        fields.Add("sceneItemId", GetSceneItemId(CurrentScene.ToString(), hidesource));
                        fields.Add("sceneItemEnabled", false);
                        _obs.SendRequest("SetSceneItemEnabled", fields);
                    }
                }
                if (showsource != "") {
                    if (showsource.Contains("/")) {
                        string[] tmp = showsource.Split("/");

                        // scene/source
                        if (tmp.Count() == 2) {
                            JObject fields = new JObject();
                            fields.Add("sceneName", tmp[0]);
                            fields.Add("sceneItemId", GetSceneItemId(tmp[0], tmp[1]));
                            fields.Add("sceneItemEnabled", true);
                            _obs.SendRequest("SetSceneItemEnabled", fields);
                        }
                    } else {
                        var CurrentScene = GetCurrentProgramScene();
                        JObject fields = new JObject();
                        fields.Add("sceneName", CurrentScene);
                        fields.Add("sceneItemId", GetSceneItemId(CurrentScene.ToString(), showsource));
                        fields.Add("sceneItemEnabled", true);
                        _obs.SendRequest("SetSceneItemEnabled", fields);
                    }
                }
                if (togglesource != "") {
                    if (togglesource.Contains("/")) {
                        string[] tmp = togglesource.Split("/");

                        // scene/source
                        if (tmp.Count() == 2)
                            OBSToggleSource(tmp[1], tmp[0]);
                    } else
                        OBSToggleSource(togglesource);
                }
                if (toggleaudio != "") {
                    JObject fields = new JObject();
                    fields.Add("inputName", toggleaudio);
                    _obs.SendRequest("ToggleInputMute", fields);
                }
                if (mute != "") {
                    JObject fields = new JObject();
                    fields.Add("inputName", mute);
                    fields.Add("inputMuted", true);
                    _obs.SendRequest("SetInputMute", fields);
                }
                if (unmute != "") {
                    JObject fields = new JObject();
                    fields.Add("inputName", unmute);
                    fields.Add("inputMuted", false);
                    _obs.SendRequest("SetInputMute", fields);
                }

                if (fadeopacity != "") {
                    // source,filtername,startopacity,endopacity,[fadedelay],[fadestep]
                    string[] tmp = fadeopacity.Split(",");
                    if (tmp.Count() < 4)
                        throw new Exception("/fadeopacity is missing required parameters!");
                    if (!IsNumericOrAsterix(tmp[2]) | !IsNumericOrAsterix(tmp[3]))
                        throw new Exception("Opacity start or end value is not nummeric (0-100)!");
                    if (tmp.Count() == 4)
                        DoSlideSetting(tmp[0], tmp[1], "opacity", tmp[2], tmp[3]);
                    else if (tmp.Count() == 5) {
                        if (!Information.IsNumeric(tmp[4]))
                            throw new Exception("Delay value is not nummeric (0-x)!");
                        int intDelay = 0; int.TryParse(tmp[4], out intDelay);
                        DoSlideSetting(tmp[0], tmp[1], "opacity", tmp[2], tmp[3], intDelay);
                    } else if (tmp.Count() == 6) {
                        if (!Information.IsNumeric(tmp[4]))
                            throw new Exception("Delay value is not nummeric (0-x)!");
                        if (!Information.IsNumeric(tmp[5]))
                            throw new Exception("Fadestep value is not nummeric (1-x)!");
                        int intDelay = 0; int.TryParse(tmp[4], out intDelay);
                        int intFadestep = 0; int.TryParse(tmp[5], out intFadestep);
                        DoSlideSetting(tmp[0], tmp[1], "opacity", tmp[2], tmp[3], intDelay, intFadestep);
                    }
                }

                if (slidesetting != "") {
                    // source,filtername,settingname,startvalue,endvalue,[slidedelay],[slidestep]
                    string[] tmp = slidesetting.Split(",");
                    if (tmp.Count() < 5)
                        throw new Exception("/slideSetting is missing required parameters!");
                    if (!IsNumericOrAsterix(tmp[3]) | !IsNumericOrAsterix(tmp[4]))
                        throw new Exception("Slide start or end value is not nummeric (0-100)!");
                    if (tmp.Count() == 5)
                        DoSlideSetting(tmp[0], tmp[1], tmp[2], tmp[3], tmp[4]);
                    else if (tmp.Count() == 6) {
                        if (!Information.IsNumeric(tmp[5]))
                            throw new Exception("Delay value is not nummeric (0-x)!");
                        int intDelay = 0; int.TryParse(tmp[5], out intDelay);
                        DoSlideSetting(tmp[0], tmp[1], tmp[2], tmp[3], tmp[4], intDelay);
                    } else if (tmp.Count() == 7) {
                        if (!Information.IsNumeric(tmp[5]))
                            throw new Exception("Delay value is not nummeric (0-x)!");
                        if (!Information.IsNumeric(tmp[6]))
                            throw new Exception("Slide step value is not nummeric (1-x)!");
                        int intDelay = 0; int.TryParse(tmp[5], out intDelay);
                        int intFadestep = 0; int.TryParse(tmp[6], out intFadestep);
                        DoSlideSetting(tmp[0], tmp[1], tmp[2], tmp[3], tmp[4], intDelay, intFadestep);
                    }
                }

                if (slideasync != "") {
                    // source,filtername,settingname,startvalue,endvalue,[slidedelay],[slidestep]
                    string[] tmp = slideasync.Split(",");
                    if (tmp.Count() < 5)
                        throw new Exception("/slideSetting is missing required parameters!");
                    if (!IsNumericOrAsterix(tmp[3]) | !IsNumericOrAsterix(tmp[4]))
                        throw new Exception("Slide start or end value is not nummeric (0-100)!");
                    if (tmp.Count() == 5) {
                        AsyncSlideSettings ExecuteTask = new AsyncSlideSettings(server, password, tmp[0], tmp[1], tmp[2], tmp[3], tmp[4]);
                        await ExecuteTask.StartSlide();
                    } else if (tmp.Count() == 6) {
                        if (!Information.IsNumeric(tmp[5]))
                            throw new Exception("Delay value is not nummeric (0-x)!");
                        int intDelay = 0; int.TryParse(tmp[5], out intDelay);
                        AsyncSlideSettings ExecuteTask = new AsyncSlideSettings(server, password, tmp[0], tmp[1], tmp[2], tmp[3], tmp[4], intDelay);
                        //System.Threading.Thread t;
                        await ExecuteTask.StartSlide();

                    } else if (tmp.Count() == 7) {
                        if (!Information.IsNumeric(tmp[5]))
                            throw new Exception("Delay value is not nummeric (0-x)!");
                        if (!Information.IsNumeric(tmp[6]))
                            throw new Exception("Slide step value is not nummeric (1-x)!");
                        int intDelay = 0; int.TryParse(tmp[5], out intDelay);
                        int intFadestep = 0; int.TryParse(tmp[6], out intFadestep);
                        AsyncSlideSettings ExecuteTask = new AsyncSlideSettings(server, password, tmp[0], tmp[1], tmp[2], tmp[3], tmp[4], intDelay, intFadestep);
                        await ExecuteTask.StartSlide();
                    }
                }

                if (setvolume != "") {
                    // source,volume,[delay],[steps]
                    string[] tmp = setvolume.Split(",");
                    if (!Information.IsNumeric(tmp[1]))
                        throw new Exception("Volume value is not nummeric (0-100)!");
                    if (tmp.Count() == 2) {
                        int intVolume = 0; int.TryParse(tmp[1], out intVolume);
                        OBSSetVolume(tmp[0], intVolume);
                    } else if (tmp.Count() == 3) {
                        if (!Information.IsNumeric(tmp[2]))
                            throw new Exception("Delay value is not nummeric (5-1000)!");
                        int intVolume = 0; int.TryParse(tmp[1], out intVolume);
                        int intDelay = 0; int.TryParse(tmp[2], out intDelay);
                        OBSSetVolume(tmp[0], intVolume, intDelay);
                    } else if (tmp.Count() == 4) {
                        if (!Information.IsNumeric(tmp[2]))
                            throw new Exception("Delay value is not nummeric (5-1000)!");
                        if (!Information.IsNumeric(tmp[3]))
                            throw new Exception("Step value is not nummeric (1-99)!");
                        int intVolume = 0; int.TryParse(tmp[1], out intVolume);
                        int intDelay = 0; int.TryParse(tmp[2], out intDelay);
                        int intSteps = 0; int.TryParse(tmp[3], out intSteps);

                        OBSSetVolume(tmp[0], intVolume, intDelay, intSteps);
                    }
                }
                if (startstream == true && !_obs.GetStreamStatus().IsActive && (title != "" && description != "" ))
                    _obs.SendRequest("StartStream");
                if (stopstream == true && _obs.GetStreamStatus().IsActive)
                    _obs.SendRequest("StopStream");
                if (startrecording == true)
                    _obs.SendRequest("StartRecord");
                if (stoprecording == true)
                    _obs.SendRequest("StopRecord");
                    
                //Wait for the stream connection to establish first
                
                if (tweet && startstream && channel_url != null) {
                    //if (channel_url.Contains("citizenmedia.news")) {
                    //    channel_url = channel_url + "&title=" + HttpUtility.UrlEncode(title);
                    //}

                    if (title != "" || description != "") {
                        //Set Restream.IO Title Automation (uses global vars
                        ChromeDriver driver = SetRestreamTitle(title, description);

                        //Start the stream!
                        _obs.SendRequest("StartStream");
                        Thread.Sleep(20000); //wait 20 seconds!
                        Console.WriteLine("Waiting 20 seconds for stream to begin before grabbing URL");
                        driver.Url = channel_url;
                        Thread.Sleep(1000);
                        string TweetMsg = driver.Url.ToString();
                        driver.Quit();

                        foreach (string tokenset in twittertokens) {
                            try {
                                String[] tokens = tokenset.Replace("/tweet=", "").Split(',');
                                string twitter_consumerkey = tokens[0];
                                string twitter_consumersecret = tokens[1];
                                string twitter_authkey = tokens[2];
                                string twitter_secret = tokens[3];

                                //Tweet this now
                                var userClient = new TwitterClient(twitter_consumerkey, twitter_consumersecret, twitter_authkey, twitter_secret);
                                if (TweetMsg != "") {
                                    var newTweet = await userClient.Tweets.PublishTweetAsync(title + " " + TweetMsg);
                                    Console.SetOut(myout);
                                    Console.WriteLine("Published Tweet: " + title + " " + TweetMsg);
                                }
                            } catch (Exception e) {
                                Console.SetOut(myout);
                                Console.WriteLine("An error occured while publishing the tweet: " + e.ToString());
                            }
                        }

                    }
                }

                if (runtime > 0) {
                    //Wait for stream to end
                    WaitForStreamEnd(runtime, myout);
                    _obs.Disconnect();
                    Console.SetOut(myout);
                    Console.WriteLine("Ok");
                }

                if (setdelay > 0 & argsindex < args.Count() & argsindex > 1) {
                    Thread.Sleep((int)Math.Round(setdelay) * 1000);
                }

            } catch (Exception ex) {
                errormessage = ex.Message.ToString();
            }
            
            try {
                _obs.Disconnect();
            } catch (Exception ex) {
            }

            // Console.SetOut(myout)
            if (errormessage == "") {
                Console.Write("Ok");
            } else {
                Console.Write("Error: " + errormessage);
            }
        }

        private static ChromeDriver SetRestreamTitle(string title, string description) {

            ChromeOptions options = new ChromeOptions();
            string localapp = AppDataFolder();
            string ProfileFolder = Path.Combine(localapp, @"Google\Chrome\User Data");
            options.AddArgument("--user-data-dir=" + ProfileFolder);
            ChromeDriver driver = new ChromeDriver(Environment.CurrentDirectory, options);

                if (showdate) {
                    title = title + " | " + DateTime.Now.Date.ToString("D", CultureInfo.CreateSpecificCulture("en-US"));
                }
                if (hashtag != "") {
                    title = title + " " + hashtag;
                }

                driver.Url = "https://app.restream.io/titles";
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                Func<IWebDriver, IWebElement> waitForTitle = new Func<IWebDriver, IWebElement>((IWebDriver Web) =>
                {
                    System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> results = Web.FindElements(By.XPath("//input[@placeholder='Title']"));
                    if (results.Count > 0) {
                        return results[0];
                    }
                    return null;
                });
                wait.Until(waitForTitle);

                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;

                //Set Title
                if (title.Trim().Length > 0) {
                    string titleClassId = driver.FindElementByXPath("//input[@placeholder='Title']").GetAttribute("class");
                    js.ExecuteScript(getReactJs(titleClassId, title));
                }

                //Set Description
                if (description.Trim().Length > 0) {
                    string descriptionClassId = driver.FindElementByXPath("//textarea[@placeholder='Description']").GetAttribute("class");
                    js.ExecuteScript(getReactJs(descriptionClassId, description));
                }

                //Get the Update All button
                string submitClassId = driver.FindElementByXPath("//button/div[text() = 'Update All']").GetAttribute("class");

                //Wait for the Update All button to be enabled...
                Func<IWebDriver, Boolean> waitForElement2 = new Func<IWebDriver, Boolean>((IWebDriver Web) =>
                {
                    Console.WriteLine("Waiting for update to save");
                    Boolean isDisabled = (Boolean)js.ExecuteScript("return document.getElementsByClassName('" + submitClassId + "')[0].hasAttribute('disabled');");
                    //System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> results = Web.FindElements(By.XPath("//*[contains(.,'Titles updated successfully')]"));
                    if (isDisabled == false) {
                        return true;
                    }
                    return false;
                });
                wait.Until(waitForElement2);

                Thread.Sleep(1000); //See if this cound be improved so this is no longer needed.

                //Click the Update All button
                js.ExecuteScript("document.getElementsByClassName('" + submitClassId + "')[0].click();");

                Thread.Sleep(1000); //See if this cound be improved so this is no longer needed.

            //Wait for the success notification to appear
            //Func<IWebDriver, IWebElement> waitForElement = new Func<IWebDriver, IWebElement>((IWebDriver Web) =>
            //    {
            //        Console.WriteLine("Waiting for update to save");
            //        System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> results = Web.FindElements(By.XPath("//*[contains(.,'Titles updated successfully')]"));
            //        if (results.Count > 0) {
            //            return results[0];
            //        }
            //        return null;
            //    });
            //    IWebElement targetElement2 = wait.Until(waitForElement);

            return driver;
        }

        private static void Connect(object sender, EventArgs e) {
            Debug.WriteLine("Connection established!");
            isConnected = true;
        }

        private static void Disconnect(object sender, OBSWebsocketDotNet.Communication.ObsDisconnectionInfo e) {
            Debug.WriteLine("Connection terminated: " + e.DisconnectReason);
            disconnectInfo = e.DisconnectReason;
        }

        private static bool IsNumericOrAsterix(string value) {
            if (value == "*")
                return true;

            if (!Information.IsNumeric(value))
                return false;

            return true;
        }

        private static JToken ConvertToType(string text) {
            if (Information.IsNumeric(text)) {
                if (text.Contains("."))
                    return double.Parse(text, System.Globalization.CultureInfo.InvariantCulture);
                else {
                    if (System.Convert.ToInt64(text) > int.MaxValue | System.Convert.ToInt64(text) < int.MinValue)
                        return System.Convert.ToInt64(text);
                    return System.Convert.ToInt32(text);
                }
            } else if (text.ToUpper() == "TRUE" | text.ToUpper() == "FALSE")
                return System.Convert.ToBoolean(text);
            else
                return text;
        }

        private static int GetSceneItemId(string sceneName, string sourceName) {
            JObject fields = new JObject();
            fields.Add("sceneName", ConvertToType(sceneName));
            fields.Add("sourceName", ConvertToType(sourceName));
            JObject response = _obs.SendRequest("GetSceneItemId", fields);
            JToken result = response.GetValue("sceneItemId");
            return ((int)result);
        }

        private static JToken GetCurrentProgramScene() {
            JObject response = _obs.SendRequest("GetCurrentProgramScene");

            return response.GetValue("currentProgramSceneName");
        }

        private static void OBSToggleSource(string source, string sceneName = "") {
            if (sceneName == "")
                sceneName = GetCurrentProgramScene().ToString();

            JObject fields = new JObject();
            fields.Add("sceneName", sceneName);
            fields.Add("sceneItemId", GetSceneItemId(sceneName, source));
            JObject sceneItemEnabled = _obs.SendRequest("GetSceneItemEnabled", fields);

            fields.Add("sceneItemEnabled", !bool.Parse(sceneItemEnabled.GetValue("sceneItemEnabled").ToString()));
            _obs.SendRequest("SetSceneItemEnabled", fields);
        }

        private static void DoSlideSetting(string source, string filtername, string settingname, string fadestart_str, string fadeend_str, int delay = 0, int fadestep = 1) {
            if (delay < 5)
                delay = 5;
            if (delay > 1000)
                delay = 1000;

            if (fadestart_str == "*" | fadeend_str == "*") {
                JObject tmpfield = new JObject();
                tmpfield.Add("sourceName", source);
                tmpfield.Add("filterName", filtername);
                JObject result = _obs.SendRequest("GetSourceFilter", tmpfield);
                if (fadestart_str == "*") {
                    JObject tmp = (JObject)result.GetValue("filterSettings");
                    fadestart_str = tmp.GetValue(settingname).ToString();
                } else if (fadeend_str == "*") {
                    JObject tmp = (JObject)result.GetValue("filterSettings");
                    fadeend_str = tmp.GetValue(settingname).ToString();
                }
            }

            bool haddecimals = false;

            int fadestart = 0; int.TryParse(fadestart_str, out fadestart);
            int fadeend = 0; int.TryParse(fadeend_str, out fadeend);

            if (fadestep < 1) {
                haddecimals = true;
                fadestart *= 100;
                fadeend *= 100;
                fadestep *= 100;
                delay /= (int)100;
            }

            JObject fields;
            if (fadestart < fadeend) {
                for (int a = fadestart; a <= fadeend; a += fadestep) {
                    fields = new JObject();
                    fields.Add("sourceName", source);
                    fields.Add("filterName", filtername);
                    JObject tmpfield = new JObject();
                    if (haddecimals == true)
                        tmpfield.Add(settingname, ConvertToType((a / (double)100).ToString()));
                    else
                        tmpfield.Add(settingname, ConvertToType(a.ToString()));
                    fields.Add("filterSettings", tmpfield);
                    _obs.SendRequest("SetSourceFilterSettings", fields);
                    System.Threading.Thread.Sleep(delay);
                }
            } else if (fadestart > fadeend) {
                for (int a = fadestart; a >= fadeend; a += -fadestep) {
                    fields = new JObject();
                    fields.Add("sourceName", source);
                    fields.Add("filterName", filtername);
                    JObject tmpfield = new JObject();
                    if (haddecimals == true)
                        tmpfield.Add(settingname, ConvertToType((a / (double)100).ToString()));
                    else
                        tmpfield.Add(settingname, ConvertToType(a.ToString()));
                    fields.Add("filterSettings", tmpfield);
                    _obs.SendRequest("SetSourceFilterSettings", fields);
                    System.Threading.Thread.Sleep(delay);
                }
            }
        }

        private static void OBSSetVolume(string source, int volume, int delay = 0, int steps = 1) {
            volume += 1;
            if (steps < 1)
                steps = 1;
            if (steps > 99)
                steps = 99;
            if (delay == 0) {
                double molvol = Math.Pow(volume, 3) / 1000000; // Convert percent to amplitude/mul (approximate, mul is non-linear)
                JObject fields = new JObject();
                fields.Add("inputName", source);
                fields.Add("inputVolumeMul", molvol);
                _obs.SendRequest("SetInputVolume", fields);
            } else {
                if (delay < 5)
                    delay = 5;
                if (delay > 1000)
                    delay = 1000;
                JObject fields = new JObject();
                fields.Add("inputName", source);
                JObject _VolumeInfo = _obs.SendRequest("GetInputVolume", fields);
                double inputVolumeMul; inputVolumeMul = System.Convert.ToDouble(_VolumeInfo.GetValue("inputVolumeMul"));
                double startvolume = Math.Pow(inputVolumeMul, 3) * 100; // Convert amplitude/mul to percent (approximate, mul is non-linear)

                if (startvolume == volume)
                    return;
                else if (startvolume < volume) {
                    for (var a = startvolume; a <= volume; a += steps) {
                        fields = new JObject();
                        fields.Add("inputName", source);
                        fields.Add("inputVolumeMul", System.Convert.ToDouble(Math.Pow(a, 3) / 1000000));
                        _obs.SendRequest("SetInputVolume", fields);
                        System.Threading.Thread.Sleep(delay);
                    }
                } else if (startvolume > volume) {
                    for (var a = startvolume; a >= volume; a += -steps) {
                        fields = new JObject();
                        fields.Add("inputName", source);
                        fields.Add("inputVolumeMul", System.Convert.ToDouble(Math.Pow(a, 3) / 1000000));
                        _obs.SendRequest("SetInputVolume", fields);
                        System.Threading.Thread.Sleep(delay);
                    }
                }
            }
        }

        private static void PrintUsage() {
            List<string> @out = new List<string>();

            @out.Add("OBSCommand v1.6.3 (for OBS Version 28.x.x and above / Websocket 5.x.x and above) ©2018-2022 by FSC-SOFT (http://www.VoiceMacro.net)");
            @out.Add(Constants.vbCrLf);
            @out.Add("Usage:");
            @out.Add("------");
            @out.Add("OBSCommand /server=127.0.0.1:4455 /password=xxxx /delay=0.5 /setdelay=0.05 /profile=myprofile /scene=myscene /hidesource=myscene/mysource /showsource=myscene/mysource /togglesource=myscene/mysource /toggleaudio=myaudio /mute=myaudio /unmute=myaudio /setvolume=mysource,volume,[delay],[steps] /fadeopacity=mysource,myfiltername,startopacity,endopacity,[fadedelay],[fadestep] /slidesetting=mysource,myfiltername,startvalue,endvalue,[slidedelay],[slidestep] /slideasync=mysource,myfiltername,startvalue,endvalue,[slidedelay],[slidestep] /startstream /stopstream /startrecording /stoprecording /command=mycommand,myparam1=myvalue1... /sendjson=jsonstring");
            @out.Add(Constants.vbCrLf);
            @out.Add("Note: If Server is omitted, default 127.0.0.1:4455 will be used.");
            @out.Add("Use quotes if your item name includes spaces.");
            @out.Add("Password can be empty if no password is set in OBS Studio.");
            @out.Add("You can use the same option multiple times.");
            @out.Add("If you use Server and Password, those must be the first 2 options!");
            @out.Add(Constants.vbCrLf);
            @out.Add("This tool uses the obs-websocket plugin to talk to OBS Studio:");
            @out.Add("https://github.com/Palakis/obs-websocket/releases");
            @out.Add(Constants.vbCrLf);
            @out.Add("3rd Party Dynamic link libraries used:");
            @out.Add("Json.NET ©2021 by James Newton-King");
            @out.Add("websocket-sharp ©2010-2022 by BarRaider");
            @out.Add("obs-websocket-dotnet ©2022 by Stéphane Lepin.");
            @out.Add(Constants.vbCrLf);
            @out.Add("Examples:");
            @out.Add("---------");
            @out.Add("OBSCommand /scene=myscene");
            @out.Add("OBSCommand /toggleaudio=\"Desktop Audio\"");
            @out.Add("OBSCommand /mute=myAudioSource");
            @out.Add("OBSCommand /unmute=\"my Audio Source\"");
            @out.Add("OBSCommand /setvolume=Mic/Aux,0,50,2");
            @out.Add("OBSCommand /setvolume=Mic/Aux,100");
            @out.Add("OBSCommand /fadeopacity=Mysource,myfiltername,0,100,5,2");
            @out.Add("OBSCommand /slidesetting=Mysource,myfiltername,contrast,-2,0,100,0.01");
            @out.Add("OBSCommand /slideasync=Mysource,myfiltername,saturation,*,5,100,0.1");
            @out.Add("OBSCommand /stopstream");
            @out.Add("OBSCommand /profile=myprofile /scene=myscene /showsource=mysource");
            @out.Add("OBSCommand /showsource=mysource");
            @out.Add("OBSCommand /hidesource=myscene/mysource");
            @out.Add("OBSCommand /togglesource=myscene/mysource");
            @out.Add("OBSCommand /showsource=\"my scene\"/\"my source\"");
            @out.Add("");
            @out.Add("For most of other simpler requests, use the generalized '/command' feature (see syntax below):");
            @out.Add("OBSCommand /command=SaveReplayBuffer");
            @out.Add(@"OBSCommand /command=SaveSourceScreenshot,sourceName=MyScene,imageFormat=png,imageFilePath=C:\OBSTest.png");
            @out.Add("OBSCommand /command=SetSourceFilterSettings,sourceName=\"Color Correction\",filterName=Opacity,filterSettings=opacity=10");
            @out.Add("OBSCommand /command=SetInputSettings,inputName=\"Browser\",inputSettings=url='https://www.google.com/search?q=query+goes+there'");
            @out.Add("");
            @out.Add("For more complex requests, use the generalized '/sendjson' feature:");
            @out.Add(@"OBSCommand.exe /sendjson=SaveSourceScreenshot={'sourceName':'MyScource','imageFormat':'png','imageFilePath':'H:\\OBSScreenShot.png'}");
            @out.Add("");
            @out.Add("You can combine multiple commands like this:");
            @out.Add("OBSCommand /scene=mysource1 /delay=1.555 /scene=mysource2 ...etc");
            @out.Add("OBSCommand /setdelay=1.555 /scene=mysource1 /scene=mysource2 ...etc");
            @out.Add(Constants.vbCrLf);
            @out.Add("Options:");
            @out.Add("--------");
            @out.Add("/server=127.0.0.1:4455            define server address and port");
            @out.Add("  Note: If Server is omitted, default 127.0.0.1:4455 will be used.");
            @out.Add("/password=xxxx                    define password (can be omitted)");
            @out.Add("/delay=n.nnn                      delay in seconds (0.001 = 1 ms)");
            @out.Add("/setdelay=n.nnn                   global delay in seconds (0.001 = 1 ms)");
            @out.Add("                                  (set it to 0 to cancel it)");
            @out.Add("/profile=myprofile                switch to profile \"myprofile\"");
            @out.Add("/scene=myscene                    switch to scene \"myscene\"");
            @out.Add("/hidesource=myscene/mysource      hide source \"scene/mysource\"");
            @out.Add("/showsource=myscene/mysource      show source \"scene/mysource\"");
            @out.Add("/togglesource=myscene/mysource    toggle source \"scene/mysource\"");
            @out.Add("  Note:  if scene is omitted, current scene is used");
            @out.Add("/toggleaudio=myaudio              toggle mute from audio source \"myaudio\"");
            @out.Add("/mute=myaudio                     mute audio source \"myaudio\"");
            @out.Add("/unmute=myaudio                   unmute audio source \"myaudio\"");
            @out.Add("/setvolume=myaudio,volume,delay,  set volume of audio source \"myaudio\"");
            @out.Add("steps                             volume is 0-100, delay is in milliseconds");
            @out.Add("                                  between steps (min. 5, max. 1000) for fading");
            @out.Add("                                  steps is (1-99), default step is 1");
            @out.Add("  Note:  if delay is omitted volume is set instant");
            @out.Add("/fadeopacity=mysource,myfiltername,startopacity,endopacity,[fadedelay],[fadestep]");
            @out.Add("                                  start/end opacity is 0-100, 0=fully transparent");
            @out.Add("                                  delay is in milliseconds, step 0-100");
            @out.Add("             Note: Use * for start- or endopacity for fade from/to current value");
            @out.Add("/slidesetting=mysource,myfiltername,settingname,startvalue,endvalue,[slidedelay],[slidestep]");
            @out.Add("                                  start/end value min/max depends on setting!");
            @out.Add("                                  delay is in milliseconds");
            @out.Add("                                  step depends on setting (can be x Or 0.x Or 0.0x)");
            @out.Add("             Note: Use * for start- or end value to slide from/to current value");
            @out.Add("/slideasync");
            @out.Add("            The same as slidesetting, only this one runs asynchron!");
            @out.Add("/startstream                      starts streaming");
            @out.Add("/stopstream                       stop streaming");
            @out.Add("/startrecording                   starts recording");
            @out.Add("/stoprecording                    stops recording");
            @out.Add("");
            @out.Add("General User Command syntax:");
            @out.Add("----------------------------");
            @out.Add("/command=mycommand,myparam1=myvalue1,myparam2=myvalue2...");
            @out.Add("                                  issues user command,parameter(s) (optional)");
            @out.Add("/command=mycommand,myparam1=myvalue1,myparam2=myvalue2,myparam3=mysubparam=mysubparamvalue");
            @out.Add("                                  issues user command,parameters and sub-parameters");
            @out.Add("");
            @out.Add("A full list of commands is available here https://github.com/obsproject/obs-websocket/blob/master/docs/generated/protocol.md");
            @out.Add("");

            int i = 0;
            double z = 0;

            while (true) {
                Console.WriteLine(@out[i]);
                if (z > Console.WindowHeight - 1) {
                    Console.Write("Press any key to continue...");
                    Console.ReadKey();
                    ClearCurrentConsoleLine();
                    z = 0;
                }
                i += 1;
                z += 1;
                if (i >= @out.Count)
                    break;

                z += @out[i].Length / (Console.WindowWidth / (double)2);
            }
        }

        private static string[] SplitWhilePreservingQuotedValues(string value, char delimiter, bool DeleteQuotes = false) {
            Regex csvPreservingQuotedStrings = new Regex(string.Format("(\"[^\"]*\"|[^{0}])+", delimiter));
            var values = csvPreservingQuotedStrings.Matches(value).Cast<Match>().Select(m => m.Value.TrimStart(' ')).Where(v => !string.IsNullOrEmpty(v));

            string[] tmp = values.ToArray();

            if (DeleteQuotes == false)
                return tmp;

            for (var a = 0; a <= tmp.Length - 1; a++) {
                if (tmp[a] != Strings.Chr(34).ToString() && tmp[a].StartsWith(Strings.Chr(34)) & tmp[a].EndsWith(Strings.Chr(34)))
                    tmp[a] = tmp[a].Substring(1, tmp[a].Length - 2);
            }

            return tmp;
        }

        public static void ClearCurrentConsoleLine() {
            int currentLineCursor = Console.CursorTop;
            Console.SetCursorPosition(0, Console.CursorTop);
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, currentLineCursor);
        }

        static void WaitForStreamEnd(double runtime, TextWriter myout) {
            if (runtime > 0) {
                Console.SetOut(myout);
                Console.WriteLine("Waiting " + (runtime / 60).ToString() + " minutes to stop streaming");
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(runtime));
                var streamStatus = _obs.GetStreamStatus();
                if (streamStatus.IsActive) {
                    try {
                        _obs.SetCurrentSceneCollection("Blank");
                    } catch {
                        //pass
                    }
                    _obs.StopStream();
                }
            }
        }
        
        public static String getReactJs(string className, string settext) {
            String injectJs = @"function setNativeValue(element, value) {
                const valueSetter = Object.getOwnPropertyDescriptor(element, 'value').set;
                const prototype = Object.getPrototypeOf(element);
                const prototypeValueSetter = Object.getOwnPropertyDescriptor(prototype, 'value').set;

                if (valueSetter && valueSetter !== prototypeValueSetter) {
                    prototypeValueSetter.call(element, value);
                } else {
                    valueSetter.call(element, value);
                }
            }
            setNativeValue(document.getElementsByClassName('" + className + "')[0], '" + HttpUtility.JavaScriptStringEncode(settext) + "'); document.getElementsByClassName('" + className + "')[0].dispatchEvent(new Event('input', { bubbles: true }));";
            return injectJs;
        }
        public static string AppDataFolder() {
            var userPath = Environment.GetEnvironmentVariable(
              RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ?
              "LOCALAPPDATA" : "Home");

            var assy = System.Reflection.Assembly.GetEntryAssembly();
            var companyName = assy.GetCustomAttributes<AssemblyCompanyAttribute>()
              .FirstOrDefault();
            var path = System.IO.Path.Combine(userPath, companyName.Company);

            return Path.GetFullPath(Path.Combine(path, @"..\"));
        }

        // Class for Async Slide Settings
        class AsyncSlideSettings {
            private string _server;
            private string _password;
            private string _source;
            private string _filtername;
            private string _settingname;
            private string _fadestart;
            private string _fadeend;
            private int _delay;
            private int _fadestep;

            public AsyncSlideSettings(string server, string password, string source, string filtername, string settingname, string fadestart, string fadeend, int delay = 0, int fadestep = 1) {
                _server = server;
                _password = password;
                _source = source;
                _filtername = filtername;
                _settingname = settingname;
                _fadestart = fadestart;
                _fadeend = fadeend;
                _delay = delay;
                _fadestep = fadestep;
            }

            async public Task StartSlide() {
                await SlideSetting(_server, _password, _source, _filtername, _settingname, _fadestart, _fadeend, _delay, _fadestep);
            }

            async public Task SlideSetting(string server, string password, string source, string filtername, string settingname, string fadestart_str, string fadeend_str, double delay = 0, int fadestep = 1) {
                var obs = new OBSWebsocket();
                obs.WSTimeout = new TimeSpan(0, 0, 0, 3);
                obs.ConnectAsync(server, password);
                int i = 0;
                while (!obs.IsConnected) {
                    System.Threading.Thread.Sleep(10);
                    i += 1;
                    if (i > 300) {
                        Console.Write("Error: can't connect to OBS websocket plugin!");
                        System.Environment.Exit(0);
                    }
                }

                if (delay < 5)
                    delay = 5;
                if (delay > 1000)
                    delay = 1000;

                if (fadestart_str == "*" | fadeend_str == "*") {
                    JObject tmpfield = new JObject();
                    tmpfield.Add("sourceName", source);
                    tmpfield.Add("filterName", filtername);
                    JObject result = obs.SendRequest("GetSourceFilter", tmpfield);

                    if (fadestart_str == "*") {
                        JObject tmp = (JObject)result.GetValue("filterSettings");
                        fadestart_str = tmp.GetValue(settingname).ToString();
                    } else if (fadeend_str == "*") {
                        JObject tmp = (JObject)result.GetValue("filterSettings");
                        fadeend_str = tmp.GetValue(settingname).ToString();
                    }
                }

                bool haddecimals = false;


                int fadestart = 0; int.TryParse(fadestart_str, out fadestart);
                int fadeend = 0; int.TryParse(fadeend_str, out fadeend);

                if (fadestep < 1) {
                    haddecimals = true;
                    fadestart *= 100;
                    fadeend *= 100;
                    fadestep *= 100;
                    delay /= (double)100;
                }

                JObject fields;
                if (fadestart < fadeend) {
                    for (int a = fadestart; a <= fadeend; a += fadestep) {
                        fields = new JObject();
                        fields.Add("sourceName", source);
                        fields.Add("filterName", filtername);
                        JObject tmpfield = new JObject();
                        if (haddecimals == true)
                            tmpfield.Add(settingname, ConvertToType((a / (double)100).ToString()));
                        else
                            tmpfield.Add(settingname, ConvertToType(a.ToString()));
                        fields.Add("filterSettings", tmpfield);
                        obs.SendRequest("SetSourceFilterSettings", fields);
                        System.Threading.Thread.Sleep((int)Math.Round(delay));
                    }
                } else if (fadestart > fadeend) {
                    for (int a = fadestart; a >= fadeend; a += -fadestep) {
                        fields = new JObject();
                        fields.Add("sourceName", source);
                        fields.Add("filterName", filtername);
                        JObject tmpfield = new JObject();
                        if (haddecimals == true)
                            tmpfield.Add(settingname, ConvertToType((a / (double)100).ToString()));
                        else
                            tmpfield.Add(settingname, ConvertToType(a.ToString()));
                        fields.Add("filterSettings", tmpfield);
                        obs.SendRequest("SetSourceFilterSettings", fields);
                        System.Threading.Thread.Sleep((int)Math.Round(delay));
                    }
                }

                _obs.Disconnect();
            }
        }
    }
}