//The MIT License(MIT)

//Copyright(c) 2016 kv1dr

//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:

//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.

//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.

using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Security.Principal;
using System.Windows.Forms;

namespace tsvOdpirac
{
    static class BasicUtils
    {
        public static bool odstraniRegisterKljuc(string mapa,string kljuc, out string napaka)
        {
            napaka = null;
            try
            {

                using (RegistryKey keys = Registry.CurrentUser.OpenSubKey(mapa, true))
                {
                    if (keys == null)
                    {
                        // Key doesn't exist. Do whatever you want to handle
                        // this case
                    }
                    else
                    {
                        if (keys.OpenSubKey(kljuc) != null)
                        {
                            RegistryKey registryKey = keys.OpenSubKey(kljuc, true);
                            foreach (string key in registryKey.GetSubKeyNames())
                                registryKey.DeleteSubKey(key);
                            registryKey.Close();
                            keys.DeleteSubKeyTree(kljuc);
                        }
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                napaka = ex.Message;
                return false;
            }
        }

        public static bool IsRunningAsAdministratior()
        {
            try
            {
                WindowsIdentity user = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(user);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool ubijProcese(string process,out string napaka)
        {
            napaka = null;
            try
            {
                foreach (Process proc in Process.GetProcessesByName(process))
                {
                    proc.Kill();
                }
                return true;
            }
            catch (Exception ex)
            {
                napaka = ex.Message;
                return false;
            }
        }

        public static void prikaziNapako(string napaka)
        {
            MessageBox.Show(napaka, "Napaka");
        }

        public static bool zazeniCmd(string ukaz, out string napaka)
        {
            napaka = null;
            try
            {
                Process process = new Process();
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                startInfo.FileName = "cmd.exe";
                startInfo.Arguments = string.Format("/C {0}",ukaz);
                process.StartInfo = startInfo;
                process.Start();
                process.WaitForExit();
                return true;
            }
            catch (Exception ex)
            {
                napaka = ex.Message;
                return false;
            }
        }

        public static bool aplikacijaNamescena(string imeAplikacije, out string napaka)
        {
            napaka = null;
            try
            {
                string registry_key = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(registry_key))
                {
                    foreach (string subkey_name in key.GetSubKeyNames())
                    {
                        using (RegistryKey subkey = key.OpenSubKey(subkey_name))
                        {
                        var display_name = subkey.GetValue("DisplayName");
                            if (display_name == null)
                                continue;
                            if (display_name.ToString().Contains(imeAplikacije))
                                return true;
                        }
                    }
                }
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(registry_key))
                {
                    foreach (string subkey_name in key.GetSubKeyNames())
                    {
                        using (RegistryKey subkey = key.OpenSubKey(subkey_name))
                        {
                            var display_name = subkey.GetValue("DisplayName");
                            if (display_name == null)
                                continue;
                            if (display_name.ToString().Contains(imeAplikacije))
                                return true;
                        }
                    }
                }

                if (!Environment.Is64BitOperatingSystem) return false;  // če ni 64-bitni OS, ne nadaljuj

                registry_key = @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall";
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(registry_key))
                {
                    foreach (string subkey_name in key.GetSubKeyNames())
                    {
                        using (RegistryKey subkey = key.OpenSubKey(subkey_name))
                        {
                            var display_name = subkey.GetValue("DisplayName");
                            if (display_name == null)
                                continue;
                            if (display_name.ToString().Contains(imeAplikacije))
                                return true;
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                napaka = ex.Message;
                return false;
            }
        }
    }
}
