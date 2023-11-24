using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WindowsInput;
using WindowsInput.Native;
using AT_ER_OutputFiles.Systems_OutputFiles.Drivers;

namespace AT_ER_OutputFiles.Systems_OutputFiles.Pages
{
    internal class LoginPage_WR : Driver
    {
        public void LoginPage(string uname, string pwd, string sytem, string port, string client)
        {
            string url = $"https://{sytem}.spinifexit.com:{port}/sap/bc/bsp/spin/erv2/webcontent/index.html?sap-client={client}&sap-user={uname}&sap-language=E&sap-sessioncmd=open";

            _driver.Navigate().GoToUrl(url);
            try
            {
                _driver.Manage().Window.Maximize();
            }
            catch (Exception)
            {
                string test = $"https://{uname}:{pwd}@{sytem}.spinifexit.com:{port}/sap/bc/bsp/spin/erv2/webcontent/index.html?sap-client={client}&sap-user={uname}&sap-language=E&sap-sessioncmd=open";
                _driver.Navigate().GoToUrl(url);
            }
            Thread.Sleep(1000);
            InputSimulator sim = new InputSimulator();
            sim.Keyboard.TextEntry($"{uname}");
            Thread.Sleep(1000);
            sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
            sim.Keyboard.TextEntry($"{pwd}");
            Thread.Sleep(1000);
            sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        }
    }
}
