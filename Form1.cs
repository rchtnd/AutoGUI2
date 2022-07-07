using Microsoft.Office.Interop.Excel;
using System;
using System.Threading;
using System.Diagnostics;
using InputSimulatorStandard;
using AutoGUI2.Properties;

namespace AutoGUI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            textBox5.Text = Settings.Default["X"].ToString();
            textBox6.Text = Settings.Default["Y"].ToString();
            textBox1.Text = Settings.Default["FilePath"].ToString();
            textBox2.Text = Settings.Default["WorkNum"].ToString();
            textBox3.Text = Settings.Default["StartRow"].ToString();
            textBox4.Text = Settings.Default["StopRow"].ToString();
        }
        #region Methods
        public string[] GetRowExcel(int x, int y, string z)
        {
            
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = excel.Workbooks.Open(z);
            Worksheet ws = wb.Worksheets[y];
            Microsoft.Office.Interop.Excel.Range row = ws.Rows[x];

            string[] arr = new string[16];
            int i = 0;
            foreach (var cell in row.Value)
            {
                arr[i] = Convert.ToString(cell);
                i++;
                if (i == 16) break;
            }
            wb.Close(0);
            excel.Quit();
            return arr;
        }
        public void CleanUp()
        {
            do
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            while (System.Runtime.InteropServices.Marshal.AreComObjectsAvailableForCleanup());
        }
        public int[] Resolution()
        {
            int[] xy = new int[2];
            xy[0] = Convert.ToInt32(textBox5.Text) - 1;
            xy[1] = Convert.ToInt32(textBox6.Text) - 1;
            return xy;
        }
        public double ConvertX(int x)
        {
            return x * 65535 / Resolution()[0];
        }
        public double ConvertY(int y)
        {
            return y * 65535 / Resolution()[1];
        }
        public void PriorityGroup(string pg)
        {
            InputSimulator Simulate = new InputSimulator();
            Simulate.Mouse.MoveMouseTo(ConvertX(226), ConvertY(187));
            Thread.Sleep(200);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(200);

            string prgr = pg.ToUpper();
            switch (prgr)
            {
                case "ROPP1":
                    Simulate.Mouse.MoveMouseTo(ConvertX(235), ConvertY(650));
                    Thread.Sleep(100);
                    Simulate.Mouse.VerticalScroll(-24);
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    Thread.Sleep(250);
                    break;
                case "ROPP2":
                    Simulate.Mouse.MoveMouseTo(ConvertX(235), ConvertY(700));
                    Thread.Sleep(100);
                    Simulate.Mouse.VerticalScroll(-24);
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    Thread.Sleep(250);
                    break;
                /*case "A3ROPP1":
                    Simulate.Mouse.MoveMouseTo(ConvertX(235), ConvertY(508));
                    Thread.Sleep(100);
                    Simulate.Mouse.VerticalScroll(-24);
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    Thread.Sleep(250);
                    break;
                case "A3ROPP2":
                    Simulate.Mouse.MoveMouseTo(ConvertX(235), ConvertY(559));
                    Thread.Sleep(100);
                    Simulate.Mouse.VerticalScroll(-24);
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    Thread.Sleep(250);
                    break;*/
                case "A1":
                    Simulate.Mouse.MoveMouseTo(ConvertX(143), ConvertY(510));
                    Simulate.Mouse.LeftButtonClick();
                    Thread.Sleep(250);
                    break;
                case "A2":
                    Simulate.Mouse.MoveMouseTo(ConvertX(143), ConvertY(654));
                    Simulate.Mouse.LeftButtonClick();
                    Thread.Sleep(250);
                    break;
                case "A3A":
                    Simulate.Mouse.MoveMouseTo(ConvertX(143), ConvertY(700));
                    Simulate.Mouse.LeftButtonClick();
                    Thread.Sleep(250);
                    break;
                case "A3B":
                    Simulate.Mouse.MoveMouseTo(ConvertX(143), ConvertY(691));
                    Thread.Sleep(100);
                    Simulate.Mouse.VerticalScroll(-3);
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    Thread.Sleep(250);
                    break;
                case "A4":
                    Simulate.Mouse.MoveMouseTo(ConvertX(143), ConvertY(684));
                    Thread.Sleep(100);
                    Simulate.Mouse.VerticalScroll(-6);
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    Thread.Sleep(250);
                    break;
                case "A5":
                    Simulate.Mouse.MoveMouseTo(ConvertX(143), ConvertY(684));
                    Thread.Sleep(100);
                    Simulate.Mouse.VerticalScroll(-8);
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    Thread.Sleep(250);
                    break;
                case "ROAP":
                    Simulate.Mouse.MoveMouseTo(ConvertX(143), ConvertY(606));
                    Thread.Sleep(100);
                    Simulate.Mouse.VerticalScroll(-24);
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    Thread.Sleep(250);
                    break;
            }
        }
        public void UniquePID(string ID)
        {
            InputSimulator Simulate = new InputSimulator();
            Simulate.Mouse.MoveMouseTo(ConvertX(263), ConvertY(262));
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);
            Simulate.Keyboard.TextEntry(ID);
            Thread.Sleep(200);
        }
        public void Names(string last, string first, string middle)
        {
            InputSimulator Simulate = new InputSimulator();

            // Last Name
            Simulate.Mouse.MoveMouseTo(ConvertX(180), ConvertY(419));
            Thread.Sleep(100);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);
            Simulate.Keyboard.TextEntry(last);
            Thread.Sleep(100);

            // First Name
            Simulate.Mouse.MoveMouseTo(ConvertX(180), ConvertY(476));
            Thread.Sleep(100);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);
            Simulate.Keyboard.TextEntry(first);
            Thread.Sleep(100);

            // Middle Name
            Simulate.Mouse.MoveMouseTo(ConvertX(180), ConvertY(543));
            Thread.Sleep(100);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);
            if (middle == null)
                Simulate.Keyboard.TextEntry("NONE");
            else
                Simulate.Keyboard.TextEntry(middle);
            Thread.Sleep(100);
        }
        public void Suffix(string sfx)
        {
            InputSimulator Simulate = new InputSimulator();

            if (sfx == null) return;

            Simulate.Mouse.MoveMouseTo(ConvertX(180), ConvertY(608));
            Thread.Sleep(100);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);
            Simulate.Keyboard.TextEntry(sfx);
            Thread.Sleep(100);
        }
        public void BDate(string bday)
        {
            InputSimulator Simulate = new InputSimulator();
            var parsedDate = DateTime.Parse(bday);

            Simulate.Mouse.MoveMouseTo(ConvertX(180), ConvertY(671));
            Thread.Sleep(100);
            Simulate.Mouse.LeftButtonClick();
            Simulate.Mouse.MoveMouseTo(ConvertX(423), ConvertY(536));
            Thread.Sleep(100);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);
            Simulate.Keyboard.TextEntry(parsedDate.ToString("MM/dd/yyyy"));
            Thread.Sleep(100);
            Simulate.Mouse.MoveMouseTo(ConvertX(858), ConvertY(439));
            Thread.Sleep(100);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(200);
        }
        public void Scroll(int x)
        {
            InputSimulator Simulate = new InputSimulator();
            Simulate.Mouse.VerticalScroll(x);
            Thread.Sleep(100);
        }
        public void Sex(string x)
        {
            InputSimulator Simulate = new InputSimulator();

            switch (x)
            {
                case "F":
                    Simulate.Mouse.MoveMouseTo(ConvertX(325), ConvertY(202));
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    Thread.Sleep(100);
                    break;
                default: // Male
                    break;
            }
        }
        public void Contact(string number)
        {
            InputSimulator Simulate = new InputSimulator();

            Simulate.Mouse.MoveMouseTo(ConvertX(209), ConvertY(254));
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);
            if (number == null)
                Simulate.Keyboard.TextEntry(number);
            else
                Simulate.Keyboard.TextEntry("4132316");
            Thread.Sleep(100);
        }
        public void Guardian(string guardian)
        {
            InputSimulator Simulate = new InputSimulator();

            Simulate.Mouse.MoveMouseTo(ConvertX(205), ConvertY(321));
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);
            if (guardian == null)
                Simulate.Keyboard.TextEntry("NONE, NONE");
            else
                Simulate.Keyboard.TextEntry(guardian);
            Thread.Sleep(100);
        }
        public void Address()
        {
            InputSimulator Simulate = new InputSimulator();

            // Region
            Simulate.Mouse.MoveMouseTo(ConvertX(153), ConvertY(390));
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);
            Simulate.Mouse.MoveMouseTo(ConvertX(118), ConvertY(609));
            Thread.Sleep(200);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(200);

            // Province
            Simulate.Mouse.MoveMouseTo(ConvertX(150), ConvertY(467));
            Thread.Sleep(100);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(200);
            Simulate.Mouse.MoveMouseTo(ConvertX(164), ConvertY(555));
            Thread.Sleep(100);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);

            // City
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(200);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);
        }
        public void Barangay(string brgy)
        {
            InputSimulator Simulate = new InputSimulator();
            Simulate.Mouse.MoveMouseTo(ConvertX(250), ConvertY(632));
            Thread.Sleep(100);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);
            if (brgy == null)
                Simulate.Keyboard.TextEntry("Poblacion");
            else
                Simulate.Keyboard.TextEntry(brgy);
            Thread.Sleep(100);
        }
        public void Vaccine(string vac)
        {
            InputSimulator Simulate = new InputSimulator();
            Simulate.Mouse.MoveMouseTo(ConvertX(158), ConvertY(138));
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);

            string vaccine = vac.ToLower();
            switch (vaccine)
            {
                case "sinovac":
                    Simulate.Mouse.MoveMouseTo(ConvertX(138), ConvertY(462));
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    break;
                case "az":
                    Simulate.Mouse.MoveMouseTo(ConvertX(138), ConvertY(509));
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    break;
                case "gamaleya":
                    Simulate.Mouse.MoveMouseTo(ConvertX(138), ConvertY(555));
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    break;
                case "pfizer":
                    Simulate.Mouse.MoveMouseTo(ConvertX(138), ConvertY(606));
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    break;
                case "moderna":
                    Simulate.Mouse.MoveMouseTo(ConvertX(138), ConvertY(650));
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    break;
                case "novavax":
                    Simulate.Mouse.MoveMouseTo(ConvertX(138), ConvertY(700));
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    break;
                case "jj":
                    Simulate.Mouse.MoveMouseTo(ConvertX(138), ConvertY(653));
                    Thread.Sleep(100);
                    Simulate.Mouse.VerticalScroll(-5);
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    break;
                case "sinopharm":
                    Simulate.Mouse.MoveMouseTo(ConvertX(138), ConvertY(700));
                    Thread.Sleep(100);
                    Simulate.Mouse.VerticalScroll(-5);
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    break;
            }
        }
        public void Dose(string dose)
        {
            InputSimulator Simulate = new InputSimulator();

            switch (dose)
            {
                case "2": // 2nd dose
                    Simulate.Mouse.MoveMouseTo(ConvertX(104), ConvertY(270));
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    break;
                case "3": // booster
                    Simulate.Mouse.MoveMouseTo(ConvertX(104), ConvertY(320));
                    Thread.Sleep(100);
                    Simulate.Mouse.LeftButtonClick();
                    break;
                default: // 1st dose
                    break;
            }
            Thread.Sleep(100);
        }
        public void Vaccinator(string vaxr)
        {
            InputSimulator Simulate = new InputSimulator();
            Simulate.Mouse.MoveMouseTo(ConvertX(126), ConvertY(431));
            Thread.Sleep(100);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);
            Simulate.Keyboard.TextEntry(vaxr);
            Thread.Sleep(200);
        }
        public void VacDate(string date)
        {
            InputSimulator Simulate = new InputSimulator();
            var parsedDate = DateTime.Parse(date);

            Simulate.Mouse.MoveMouseTo(ConvertX(117), ConvertY(493));
            Thread.Sleep(100);
            Simulate.Mouse.LeftButtonClick();
            Simulate.Mouse.MoveMouseTo(ConvertX(423), ConvertY(536));
            Thread.Sleep(100);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);
            Simulate.Keyboard.TextEntry(parsedDate.ToString("MM/dd/yyyy"));
            Thread.Sleep(100);
            Simulate.Mouse.MoveMouseTo(ConvertX(858), ConvertY(439));
            Thread.Sleep(100);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);
        }
        public void BatchLot(string batchlot)
        {
            InputSimulator Simulate = new InputSimulator();

            // Batch Number
            Simulate.Mouse.MoveMouseTo(ConvertX(102), ConvertY(489));
            Thread.Sleep(100);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);
            Simulate.Keyboard.TextEntry(batchlot);
            Thread.Sleep(200);

            // Lot Number
            Simulate.Mouse.MoveMouseTo(ConvertX(102), ConvertY(550));
            Thread.Sleep(100);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);
            Simulate.Keyboard.TextEntry(batchlot);
            Thread.Sleep(200);
        }
        public void Save()
        {
            InputSimulator Simulate = new InputSimulator();
            Simulate.Mouse.MoveMouseTo(ConvertX(1243), ConvertY(688));
            Thread.Sleep(100);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);
        }
        public void Start()
        {
            InputSimulator Simulate = new InputSimulator();
            Simulate.Mouse.MoveMouseTo(ConvertX(650), ConvertY(653));
            Thread.Sleep(100);
            Simulate.Mouse.LeftButtonClick();
            Thread.Sleep(100);
        }
        #endregion
        private void button1_Click(object sender, EventArgs e) // "Main"
        {
            Settings.Default["X"] = textBox5.Text;
            Settings.Default["Y"] = textBox6.Text;
            Settings.Default["FilePath"] = textBox1.Text;
            Settings.Default["WorkNum"] = textBox2.Text;
            Settings.Default["StartRow"] = textBox3.Text;
            Settings.Default["StopRow"] = textBox4.Text;
            Settings.Default.Save();

            button1.Text = "Running...";

            Stopwatch stopWatch = new Stopwatch();
            int rowStart = Convert.ToInt16(textBox3.Text);
            int rowStop = Convert.ToInt16(textBox4.Text);
            int wsNum = Convert.ToInt16(textBox2.Text);
            string filePath = textBox1.Text;

            Thread.Sleep(1000);
            stopWatch.Start();
            while (rowStart < rowStop + 1)
            {
                string[] row = GetRowExcel(rowStart, wsNum, filePath);
                CleanUp();
                Start();
                PriorityGroup(row[0]);
                UniquePID(row[1]);
                Names(row[2], row[3], row[4]);
                Suffix(row[5]);
                BDate(row[9]);
                Scroll(-30);
                Sex(row[8]);
                Contact(row[6]);

                if (row[0].ToUpper() == "ROPP1" || row[0].ToUpper() == "ROPP2")
                    Guardian(row[15]);
                else
                    Scroll(3);

                Address();
                Barangay(row[7]);
                Scroll(-33);
                Vaccine(row[14]);

                if (row[14].ToLower() == "jj") Scroll(2);

                Dose(row[11]);
                Vaccinator(row[10]);
                VacDate(row[13]);
                Scroll(-4);
                BatchLot(row[12]);
                Save();
                CleanUp();

                rowStart++;
            }

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;
            string elapsedTime = String.Format("\n{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
            textBox7.Text += "Time elapsed: " + elapsedTime;
            button1.Text = "Run";
        }
    }
}