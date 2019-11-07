using System;
using System.Runtime.InteropServices;
using System.Net;
using System.Windows.Forms;
using System.Diagnostics;
using System.Reflection;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Web.Script.Serialization;
using System.Net.Mail;



namespace sample.programs
{
    class Program
    {
        static void Main(string[] args)
        {
            SampleJson sj = new SampleJson();
            string path = @"D:\Test\SPO\02010000.xml";
            string ExportPath = @"D:\Test\"; 
            //string date = "10\\03\\2019";
            sj.CreateJSON("MIN_CF_STOCKPO", "WhareHouse Y", "AC1001", "Ship to xyz", "PO999", "10\\02\\2019", path, ExportPath,  "1234","1");
      }
    }
}
