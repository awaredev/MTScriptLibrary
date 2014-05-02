using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace MTScriptLibrary
{
    class Program
    { 
        //This is an Executable for testing the MTWrapper. It will eventually be discarded to allow compiling
        //The core MTWrapper system as a DLL to be integrated into multiple projects.
        static void Main(string[] args)
        {
            //Test Absolute to Relative Path system to work around path issues
            string sPath = @"C:\Program Files\Meditech\MagicCS\Client\VMAGIC.EXE"; //@"..\..\Program Files\Meditech\MagicCS\Client\VMAGIC.EXE";//
            MTScriptSession session = new MTScriptSession(sPath,false );
            Console.WriteLine("STARTED");
            //session.Timeout = 30;
            Console.WriteLine(session.Title);
            //Console.WriteLine(session.Timeout);
            session.SendString("clairobc",true);
            Console.WriteLine(session.Title);
            //Console.WriteLine("NAME ENTERED");
            session.SendSpecKey(MTScriptSession.SpecialKey.Tab);
            //Console.WriteLine("TAB");
            //Console.WriteLine("next");
            session.SendPassword("siobhan8");
            //Console.WriteLine("PASS ENTERED");
            session.SendSpecKey(MTScriptSession.SpecialKey.Tab);
            //Console.WriteLine("TAB");
            session.SendSpecKey(MTScriptSession.SpecialKey.Return);
            //Console.WriteLine("RETURN");

            while (session.Title != "Lookup")
            { Console.WriteLine(session.Title); } //Loop until screen is lookup
            //Console.WriteLine(session.CurrentField);
            //Console.WriteLine(session.Title);
            //Console.WriteLine(session.SendString("m",false));
            //MTScriptSession.MTWindow win = session.CurrentWindow;
            //Console.WriteLine(win.XPosition + " " + win.YPosition );
            session.SendString("m", false);
            while (session.Title != "Lookup")
            { Console.WriteLine(session.Title); }
            //Console.WriteLine(win.XPosition + " " + win.YPosition);
            //Console.ReadKey();
            //Console.WriteLine("M");
            session.SendSpecKey(MTScriptSession.SpecialKey.Return);
            //Console.WriteLine("RETURN");
            //Console.WriteLine(session.CurrentField);
            //session.Timeout = 5;
            //Console.WriteLine("FINAL SCREEN");
            Console.WriteLine(session.Title);
            Console.WriteLine(session.CurrentWindow.TabCount);
            Console.WriteLine(session.CurrentWindow.IsMessageBox);
            Console.WriteLine(session.CurrentWindow.XPosition);
            Console.WriteLine(session.CurrentField);
            //while (true)
            //{ int i = 1; }
            Console.ReadKey();
            Console.WriteLine(session.CurrentField);
            Console.ReadKey();
            session.Close();
            //string sPath = "MagicCS/Client/VMAGIC.EXE"; //
            //MTScriptSession session = new MTScriptSession(sPath);
            //Console.WriteLine("launched");
            //byte[] myByte = { (byte)100, (byte)233, (byte)38 };
            //IntPtr buff = GCHandle.ToIntPtr(GCHandle.Alloc(myByte));
            //MTWrapper.MCSSetTimeout(1);
            //Console.WriteLine(MTWrapper.MCSEnterKeys(8, buff));
            //Console.WriteLine(GCHandle.FromIntPtr(buff).Target.ToString()) ;
            //Console.ReadKey();
        }
    }
}
