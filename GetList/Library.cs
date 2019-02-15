﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetList
{
    class Library
    {
        public static void WriteLog(Exception ex)
        {
            StreamWriter sw = null;

            try
            {
                sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\ErrorLog.txt", true);
                sw.WriteLine(DateTime.Now.ToString() + " : " + ex.Source.ToString().Trim() + " ; " + ex.Message.ToString().Trim());
                if (ex.InnerException != null)
                {
                    sw.WriteLine(DateTime.Now.ToString().ToString() + " : " + ex.Source.ToString().Trim() + " ; " + Convert.ToString(ex.InnerException.Message));
                }


                sw.Flush();
                sw.Close();

            }
            catch
            {

            }
        }
        public static void WriteLog(String Message, Exception ex)
        {
            StreamWriter sw = null;

            try
            {
                sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\ErrorLog.txt", true);

                sw.WriteLine(DateTime.Now.ToString() + " : " + Message);
                sw.WriteLine(ex.Source.ToString().Trim() + " ; " + ex.Message.ToString().Trim());
                if (ex.InnerException != null)
                {
                    sw.WriteLine(DateTime.Now.ToString() + " : " + ex.Source.ToString().Trim() + " ; " + Convert.ToString(ex.InnerException.Message));
                }
                sw.Flush();
                sw.Close();

            }
            catch
            {

            }
        }
        public static void WriteLog(String Message)
        {
            StreamWriter sw = null;

            try
            {
                sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\ErrorLog.txt", true);


                if (Message != "")
                {
                    sw.WriteLine(DateTime.Now.ToString() + " : " + Message);
                    sw.WriteLine("");
                }
                else
                {
                    sw.WriteLine("");
                }
                sw.Flush();
                sw.Close();

            }
            catch
            {

            }
        }
    }
}
