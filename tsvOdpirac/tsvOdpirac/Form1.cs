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

using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace tsvOdpirac
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            bool isError = false;
            if (MessageBox.Show("Želite, da ta program nastavi odpiranje tsv datotek z Excelom?", string.Empty, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                Close();
                return;
            }
            string napaka;
            if (!BasicUtils.IsRunningAsAdministratior())
            {
                BasicUtils.prikaziNapako("Aplikacijo morate zagnati kot skrbnik!");
                return;
            }
            if (!BasicUtils.odstraniRegisterKljuc(@"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\", ".tsv", out napaka) && napaka != null)
            {
                BasicUtils.prikaziNapako(napaka);
                isError = true;
            }
            if (!BasicUtils.ubijProcese("explorer", out napaka) && napaka != null)
            {
                BasicUtils.prikaziNapako(napaka);
                isError = true;
            }
            bool excelNamescen = BasicUtils.aplikacijaNamescena("Microsoft Office", out napaka);
            if (!excelNamescen && napaka != null)
            {
                BasicUtils.prikaziNapako(napaka);
                isError = true;
            }
            bool libreOfficeNamescen = BasicUtils.aplikacijaNamescena("LibreOffice", out napaka);
            if (!excelNamescen && napaka != null)
            {
                BasicUtils.prikaziNapako(napaka);
                isError = true;
            }
            string program = "Microsoft Excel";
            string ukaz = @"assoc .tsv=Excel.Sheet.8";
            if((excelNamescen && libreOfficeNamescen) || (!excelNamescen && !libreOfficeNamescen))
            {
                MessageBoxManager.Yes = "Microsoft Excel";
                MessageBoxManager.No = "LibreOffice Calc";
                MessageBoxManager.Register();
                string izbira = "";
                if (excelNamescen && libreOfficeNamescen)
                    izbira = "Program je zaznal, da imate nameščen tako Microsoft Excel, kot LibreOffice Calc. S katerim programom želite odpirati tsv datoteke?";
                else if (!excelNamescen && !libreOfficeNamescen)
                    izbira = "Program je zaznal, da nimate nameščenega ne programa Microsoft Excel, ne LibreOffice Calc. Če imate katerega od teh dveh programov vseeno nameščene, izberite spodaj s katerim programom želite odpirati tsv datoteke.";
                DialogResult rezultat = MessageBox.Show(izbira,
                    string.Empty, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (rezultat == DialogResult.No)
                {
                    ukaz = @"assoc .tsv=LibreOffice.CalcDocument.1";
                    program = "LibreOffice Calc";
                }
                MessageBoxManager.Unregister();
            } else if(libreOfficeNamescen)
            {
                ukaz = @"assoc .tsv=LibreOffice.CalcDocument.1";
                program = "LibreOffice Calc";
            }
            if (!BasicUtils.zazeniCmd(ukaz, out napaka) && napaka != null)
                {
                BasicUtils.prikaziNapako(napaka);
                isError = true;
            }
            if (!BasicUtils.ubijProcese("explorer", out napaka) && napaka != null)
            {
                BasicUtils.prikaziNapako(napaka);
                isError = true;
            }
            Process.Start("explorer.exe");
            progressBar1.Style = ProgressBarStyle.Blocks;
            progressBar1.Value = 100;
            if (isError)
            {
                MessageBox.Show("Prosimo sporočite prej prikazano napako avtorju!");
            }
            else
            {
                MessageBox.Show(string.Format("Nastavili smo odpiranje .tsv datotek z {0}-om. Sedaj lahko z {0}-om odpirate datoteke s .tsv končnicami!",program));
            }
            Close();
        }
    }
}
