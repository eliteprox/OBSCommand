﻿using System;
using System.IO;
using System.Text;
using OBSWebsocketDotNet;
using System.Collections;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Diagnostics;

using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System.Globalization;
using OpenQA.Selenium.Support.UI;

namespace OBSCommand {
    class Program {
        private static OBSWebsocket _obs;
        static int Main(string[] args) {
            string server = "ws://127.0.0.1:4444";
            string password = "";
            string profile = "";
            string scene = "";
            string hidesource = "";
            string showsource = "";
            string toggleaudio = "";
            string mute = "";
            string prescene = "";
            string unmute = "";
            //string setvolume = "";
            bool stopstream = false;
            bool startstream = false;
            bool startrecording = false;
            bool stoprecording = false;

            double runtime = 0;
            string title = "";
            string description = "";
            bool showdate = false;
            string hashtag = "";

            if (args.Length == 0) {
                PrintUsage();
                return 0;
            }

            foreach (string arg in args) {
                if (arg.StartsWith("/server=")) {
                    server = "ws://" + arg.Replace("/server=", "");
                }
                if (arg.StartsWith("/password=")) {
                    password = arg.Replace("/password=", "");
                }
                if (arg.StartsWith("/profile=")) {
                    profile = arg.Replace("/profile=", "");
                }
                if (arg.StartsWith("/scene=")) {
                    scene = arg.Replace("/scene=", "");
                }
                if (arg.StartsWith("/prescene=")) {
                    prescene = arg.Replace("/prescene=", "");
                }
                if (arg.StartsWith("/hidesource=")) {
                    hidesource = arg.Replace("/hidesource=", "");
                }
                if (arg.StartsWith("/showsource=")) {
                    showsource = arg.Replace("/showsource=", "");
                }
                if (arg.StartsWith("/toggleaudio=")) {
                    toggleaudio = arg.Replace("/toggleaudio=", "");
                }
                if (arg.StartsWith("/mute=")) {
                    mute = arg.Replace("/mute=", "");
                }
                if (arg.StartsWith("/unmute=")) {
                    unmute = arg.Replace("/unmute=", "");
                }
                //if (arg.StartsWith("/setvolume=")) {
                //    setvolume = arg.Replace("/setvolume=", "");
                //}
                if (arg.StartsWith("/startstream")) {
                    startstream = true;
                }
                if (arg.StartsWith("/stopstream")) {
                    stopstream = false;
                }
                if (arg.StartsWith("/startrecording")) {
                    startrecording = true;
                }
                if (arg.StartsWith("/stoprecording")) {
                    stoprecording = true;
                }
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

                if (arg.StartsWith("/description=")) {
                    description = arg.Replace("/description=", "");
                }

                if (arg.StartsWith("/title=")) {
                    if (arg.Contains(",")) {
                        string[] tmp = arg.Split(',');
                        title = tmp[0].Replace("/title=", "");
                        if (tmp.Length > 1) {
                            bool.TryParse(tmp[1], out showdate);
                        }
                        if (tmp.Length > 2) {
                            hashtag = tmp[2];
                        }
                    } else {
                        title = arg.Replace("/title=", "");
                    }

                    if (arg.StartsWith("/showdate")) {
                        showdate = true;
                    }
                    if (arg.StartsWith("/hashtag")) {
                        hashtag = arg.Replace("/hashtag=", "");
                    }
                }
            }

            TextWriter myout = Console.Out;
            try {
                StringBuilder builder = new StringBuilder();
                TextWriter writer = new StringWriter(builder);

                Console.SetOut(writer);
                _obs = new OBSWebsocket();
                _obs.WSTimeout = new TimeSpan(0, 0, 0, 3);
                _obs.Connect(server, password);

                if (title != "" || description != "") {
                    if (showdate) {
                        title = title + " | " + DateTime.Now.Date.ToString("D", CultureInfo.CreateSpecificCulture("en-US"));
                    }
                    if (hashtag != "") {
                        title = title + " " + hashtag;
                    }

                    var driver = new EdgeDriver();
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
                        string descriptionClassId = driver.FindElementByXPath("//input[@placeholder='Description']").GetAttribute("class");
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

                    //Click the Update All button
                    js.ExecuteScript("document.getElementsByClassName('" + submitClassId + "')[0].click();");

                    //Wait for the success notification to appear
                    Func<IWebDriver, IWebElement> waitForElement = new Func<IWebDriver, IWebElement>((IWebDriver Web) =>
                    {
                        Console.WriteLine("Waiting for update to save");
                        System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> results = Web.FindElements(By.XPath("//*[contains(.,'Titles updated successfully')]"));
                        if (results.Count > 0) {
                            return results[0];
                        }
                        return null;
                    });
                    IWebElement targetElement2 = wait.Until(waitForElement);

                    driver.Quit();
                }

                if (profile != "") {
                    _obs.SetCurrentProfile(profile);
                }

                if (scene != "") {
                    _obs.SetCurrentScene(scene);
                }

                if (hidesource != "") {
                    if (hidesource.Contains("/")) {
                        string[] tmp = hidesource.Split("/");
                        if (tmp.Length == 2) {
                            // scene/source
                            _obs.SetSourceRender(tmp[1], false, tmp[0]);
                        }
                    } else {
                        _obs.SetSourceRender(hidesource, false);
                    }
                }

                if (showsource != "") {
                    if (showsource.Contains("/")) {
                        string[] tmp = showsource.Split("/");
                        if (tmp.Length == 2) {
                            // scene/source
                            _obs.SetSourceRender(tmp[1], false, tmp[0]);
                        }
                    } else {
                        _obs.SetSourceRender(showsource, true);
                    }
                }

                if (toggleaudio != "") {
                    _obs.ToggleMute(toggleaudio);
                }

                if (mute != "") {
                    _obs.SetMute(mute, true);
                }

                if (unmute != "") {
                    _obs.SetMute(unmute, false);
                }

                if (stopstream) {
                    _obs.StopStreaming();
                }

                if (startrecording) {
                    _obs.StartRecording();
                }

                if (stoprecording) {
                    _obs.StopRecording();
                }

                if (startstream) {
                    var streamStatus = _obs.GetStreamingStatus();
                    if (_obs.GetStreamingStatus().IsStreaming) {
                        _obs.StopStreaming();
                    }
                    while (_obs.GetStreamingStatus().IsStreaming) {
                        System.Threading.Thread.Sleep(1000);
                    }
                    _obs.StartStreaming();
                    if (runtime > 0) {
                        Console.SetOut(myout);
                        Console.WriteLine("Waiting " + (runtime / 60).ToString() + " minutes to stop streaming");
                        System.Threading.Thread.Sleep(TimeSpan.FromSeconds(runtime));
                        if (_obs.GetStreamingStatus().IsStreaming) {
                            try {
                                _obs.SetCurrentScene("Blank");
                            } catch {
                                //pass
                            }
                            _obs.StopStreaming();
                        }
                    }
                }

                _obs.Disconnect();
                Console.SetOut(myout);
                Console.WriteLine("Ok");

                return 1;
            } catch (Exception ex) {
                Console.SetOut(myout);
                Console.WriteLine("Error: " + ex.Message.ToString());
                return 0;
            }
        }
        static void PrintUsage() {
            List<string> output = new List<string>();
            output.Add("OBSCommand v1.3 ©2018 by FSC-SOFT (http://www.VoiceMacro.net);");
            output.Add(Environment.NewLine);
            output.Add("Usage:");
            output.Add("------");
            //output.Add("OBSCommand.exe /server=127.0.0.1:4444 /password=xxxx /profile=myprofile /scene=myscene /hidesource=myscene/mysource /showsource=myscene/mysource /toggleaudio=myaudio /mute=myaudio /unmute=myaudio /setvolume=mysource,volume,delay /startstream /stopstream /startrecording /stoprecording");
            output.Add("OBSCommand.exe /server=127.0.0.1:4444 /password=xxxx /profile=myprofile /scene=myscene /hidesource=myscene/mysource /showsource=myscene/mysource /toggleaudio=myaudio /mute=myaudio /unmute=myaudio /title=restreamtitle,showdate,hashtag /startstream /stopstream /startrecording /stoprecording");
            output.Add(Environment.NewLine);
            output.Add("Note: If Server is omitted, default 127.0.0.1:4444 will be used.");
            output.Add("Use quotes if your item name includes spaces.");
            output.Add("Password can be empty if no password is set in OBS Studio.");
            output.Add(Environment.NewLine);
            output.Add("This tool uses the obs-websocket plugin to talk to OBS Studio:");
            output.Add("https://github.com/Palakis/obs-websocket/releases");
            output.Add(Environment.NewLine);
            output.Add("Dynamic link libraries used:");
            output.Add("Json.NET ©2008 by James Newton-King");
            output.Add("websocket-sharp ©2010-2016 by sta.blockhead");
            output.Add("obs-websocket-dotnet ©2017 by Stéphane Lepin.");
            output.Add(Environment.NewLine);
            output.Add("Examples:");
            output.Add("---------");
            output.Add("OBSCommand.exe /scene=myscene");
            output.Add(@"OBSCommand.exe /toggleaudio=""Desktop Audio""");
            output.Add("OBSCommand.exe /mute=myAudioSource");
            output.Add(@"OBSCommand.exe /unmute=""my Audio Source""");
            //output.Add("OBSCommand.exe /setvolume=Mic/Aux,0,50");
            //output.Add("OBSCommand.exe /setvolume=Mic/Aux,100");
            output.Add("OBSCommand.exe /stopstream");
            output.Add("OBSCommand.exe /profile=myprofile /scene=myscene /showsource=mysource");
            output.Add("OBSCommand.exe /showsource=mysource");
            output.Add("OBSCommand.exe /hidesource=myscene/mysource");
            output.Add(@"OBSCommand.exe /showsource=""my scene""/""my source""");
            output.Add(@"OBSCommand.exe /title=""my title"",true/false (add date to title),""optional hashtag""");
            output.Add("OBSCommand.exe /runtime=minutes.seconds");
            output.Add(Environment.NewLine);
            output.Add("Options:");
            output.Add("--------");
            output.Add("/server=127.0.0.1:4444            define server address and port");
            output.Add("  Note: If Server is omitted, default 127.0.0.1:4444 will be used.");
            output.Add("/password=xxxx                    define password (can be omitted);");
            output.Add(@"/profile=myprofile                switch to profile ""myprofile""");
            output.Add(@"/scene=myscene                    switch to scene ""myscene""");
            output.Add(@"/hidesource=myscene/mysource      hide source ""scene/mysource""");
            output.Add(@"/showsource=myscene/mysource      show source ""scene/mysource""");
            output.Add("  Note:  if scene is omitted, current scene is used");
            output.Add(@"/toggleaudio=myaudio              toggle mute from audio source ""myaudio""");
            output.Add(@"/mute=myaudio                     mute audio source ""myaudio""");
            output.Add(@"/unmute=myaudio                   unmute audio source ""myaudio""");
            output.Add(@"/setvolume=myaudio,volume,delay   set volume of audio source ""myaudio""");
            output.Add("                                  volume is 0-100, delay is in milliseconds");
            output.Add("                                  between steps (min. 10, max. 1000); for fading");
            output.Add("  Note:  if delay is omitted volume is set instant");
            output.Add("/startstream                      starts streaming");
            output.Add("/stopstream                       stop streaming");
            output.Add("/startrecording                   starts recording");
            output.Add("/stoprecording                    stops recording");
            output.Add("");

            int i = 0;
            int z = 0;

            while (true)
            {
                Console.WriteLine(output[i]);
                if (z == Console.WindowHeight - 6) {
                    Console.Write("Press any key to continue...");
                    Console.ReadKey();
                    ClearCurrentConsoleLine();
                    z = 0;
                }
                i += 1;
                z += 1;
                if (i >= output.Count) { break; }
                if (output[i].Length > Console.WindowWidth) { z += 1; }
                if (output[i].Length > Console.WindowWidth * 2) { z += 1; }
            }


        }

        static void ClearCurrentConsoleLine() {
            int currentLineCursor = Console.CursorTop;
            Console.SetCursorPosition(0, Console.CursorTop);
            Console.Write(" " + Console.WindowWidth.ToString());
            Console.SetCursorPosition(0, currentLineCursor);
        }

        public void WaitForElement(IWebDriver webDriver, IWebElement element, int timeout = 5) {
            WebDriverWait wait = new WebDriverWait(webDriver, TimeSpan.FromMinutes(timeout));
            wait.IgnoreExceptionTypes(typeof(NoSuchElementException));
            wait.IgnoreExceptionTypes(typeof(StaleElementReferenceException));
            wait.Until<bool>(driver =>
            {
                try {
                    return element.Displayed;
                } catch (Exception) {
                    return false;
                }
            });
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
            setNativeValue(document.getElementsByClassName('" + className + "')[0], '" + settext + "'); document.getElementsByClassName('" + className + "')[0].dispatchEvent(new Event('input', { bubbles: true }));";
            return injectJs;
        }
    }
}
