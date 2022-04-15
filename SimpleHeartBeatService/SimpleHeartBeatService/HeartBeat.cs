using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace SimpleHeartBeatService
{
    public class HeartBeat
    {
        private readonly Timer timer1;
        private string timeString;
        public int getCallType;
        public HeartBeat()
        {
            //timer1 = new Timer(1000)
            //{
            //    AutoReset=true
            //};
            //timer1.Elapsed += TimerElapsed;
            int strTime = Convert.ToInt32(ConfigurationSettings.AppSettings["callDuration"]);
            getCallType = Convert.ToInt32(ConfigurationSettings.AppSettings["CallType"]);
            if (getCallType == 1)
            {
                timer1 = new System.Timers.Timer();
                double inter = (double)GetNextInterval();
                timer1.Interval = inter;
                timer1.Elapsed += new ElapsedEventHandler(ServiceTimer_Tick);
            }
            else
            {
                timer1 = new System.Timers.Timer();
                timer1.Interval = strTime * 1000;
                timer1.Elapsed += new ElapsedEventHandler(ServiceTimer_Tick);
            }
        }


        /////////////////////////////////////////////////////////////////////
         public void Start()
        {
            timer1.AutoReset = true;
            timer1.Enabled = true;
            SendMailService.writeLog("Service started" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss.ffff"));
        }

        /////////////////////////////////////////////////////////////////////
        public void Stop()
        {
            timer1.AutoReset = false;
            timer1.Enabled = false;
            SendMailService.writeLog("Service stopped" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss.ffff"));

        }

        /////////////////////////////////////////////////////////////////////
        private double GetNextInterval()
        {
            timeString = ConfigurationSettings.AppSettings["StartTime"];
            SendMailService.writeLog("timeString"+timeString );
            DateTime t = DateTime.Parse(timeString);
            TimeSpan ts = new TimeSpan();
            int x;
            ts = t - System.DateTime.Now;
            if (ts.TotalMilliseconds < 0)
            {
                ts = t.AddDays(1) - System.DateTime.Now;//Here you can increase the timer interval based on your requirments.   
            }
            return ts.TotalMilliseconds;
        }

        /////////////////////////////////////////////////////////////////////
        private void SetTimer()
        {
            try
            {
                double inter = (double)GetNextInterval();
                timer1.Interval = inter;
                timer1.Start();
            }
            catch (Exception ex)
            {
            }
        }

        /////////////////////////////////////////////////////////////////////
        private void ServiceTimer_Tick(object sender, System.Timers.ElapsedEventArgs e)
        {
            SendMailService sendMailService = new SendMailService();
            sendMailService.sendEMailThroughOUTLOOK();

            if (getCallType == 1)
            {
                timer1.Stop();
                System.Threading.Thread.Sleep(1000000);
                SetTimer();
            }
        }




        //private void TimerElapsed(object sender, ElapsedEventArgs e)
        //{
        //    string[] lines = new string[] { DateTime.Now.ToString() };
        //    File.AppendAllLines("D:\\b.test", lines);
        //}

        //public void Start()
        //{
        //    timer1.Start();
        //}
        //public void Stop()
        //{
        //    timer1.Stop();
        //}
    }
}
