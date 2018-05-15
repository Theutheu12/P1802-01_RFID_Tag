//using System;
//using System.Collections.Generic;
//using System.Windows.Forms;
//using ReaderB;

using System;
using System.Text;
using System.Threading;
using ReaderB;

namespace UHFReader18demomain
{
    static class Program
    {
        static ConsoleKeyInfo chr;

        static void Main()
        {

            int ComPortName = 0;
            byte DevAddr = 0x00;
            byte ComBaudRate = 6;
            int DevHandle = 0;
            int ret;
            byte[] fPassWord = new byte[4];
            byte[] TrType = new byte[2];
            byte[] VersionInfo = new byte[2];
            byte ReaderType = 0;
            byte ScanTime = 0;
            byte dmaxfre = 0;
            byte dminfre = 0;
            byte powerdBm = 0;
            byte AdrTID = 0;
            byte LenTID = 2;
            byte TIDFlag = 0;
            int CardNum = 0;
            int Totallen = 0;
            byte[] EPC = new byte[5000];

            DateTime scanTime = DateTime.Now;

            //string tmp_str;


            Thread Thr_EndProgramm = new Thread(EndProgramm);
            Thr_EndProgramm.Start();

            Console.WriteLine("***** VT UHF Tag Reader v0.1 *****");
            Console.WriteLine("");

            ret = StaticClassReaderB.AutoOpenComPort(ref ComPortName, ref DevAddr, ComBaudRate, ref DevHandle);

            if (ret != 0)
            {
                Console.WriteLine("Unable to detect a tag reader");
                Console.Write("Error n°: ");
                Console.WriteLine(Convert.ToString(ret));
                Console.WriteLine("Press any key to quit...");
                //Console.ReadKey();
                //return;
            }

            Console.WriteLine("Reader opened");

            ret = StaticClassReaderB.GetReaderInformation(ref DevAddr, VersionInfo, ref ReaderType, TrType, ref dmaxfre, ref dminfre, ref powerdBm, ref ScanTime, DevHandle);

            if (ret != 0)
            {

                Console.Write("Error n°: ");
                Console.WriteLine(Convert.ToString(ret));
                Console.WriteLine("Press any key to quit...");
                //Console.ReadKey();
                //return;
            }

            Console.Write("Reader version : ");
            Console.WriteLine(Convert.ToString(VersionInfo[0], 10).PadLeft(2, '0') + "." + Convert.ToString(VersionInfo[1], 10).PadLeft(2, '0'));
            Console.Write("Device adress : ");
            Console.WriteLine(Convert.ToString(DevAddr, 16).PadLeft(2, '0'));
            Console.Write("Reader power [dBm] : ");
            Console.WriteLine(Convert.ToString(powerdBm, 10).PadLeft(2, '0'));
            Console.Write("Scan time : ");
            Console.WriteLine(Convert.ToString(ScanTime, 10).PadLeft(2, '0') + "*100ms");
            //FreBand = Convert.ToByte(((dmaxfre & 0xc0) >> 4) | (dminfre >> 6));*/

            if (ret != 0)
            {

                Console.Write("Error n°: ");
                Console.WriteLine(Convert.ToString(ret));
                Console.WriteLine("Press any key to quit...");
                //Console.ReadKey();
                //return;
            }

            do
            {
                ret = StaticClassReaderB.Inventory_G2(ref DevAddr, AdrTID, LenTID, TIDFlag, EPC, ref Totallen, ref CardNum, DevHandle);

                if (ret == 1)
                {
                    byte[] daw = new byte[Totallen];
                    Array.Copy(EPC, daw, Totallen);
                    StringBuilder sb = new StringBuilder(daw.Length * 3);
                    foreach (byte b in daw)
                        sb.Append(Convert.ToString(b, 16).PadLeft(2, '0'));

                    Console.Write(DateTime.Now.TimeOfDay.ToString());
                    Console.Write(" ");
                    Console.Write(CardNum);
                    Console.Write(" ");
                    Console.Write(Totallen);
                    Console.Write(" ");

                    //Console.WriteLine(sb.ToString().ToUpper());
                    Console.Write(Convert.ToString(EPC[0], 16));
                    Console.Write(Convert.ToString(EPC[1], 16));
                    Console.Write(Convert.ToString(EPC[2], 16));
                    Console.Write(Convert.ToString(EPC[3], 16));
                    Console.WriteLine("");
                }

                Thread.Sleep(10);

            }while (chr.Key != ConsoleKey.C);
            
            Console.WriteLine("Press any key to quit...");
            Console.ReadKey();
            StaticClassReaderB.CloseComPort();
            
        }

        //La méthode prend en paramètre un et un seul paramètre de type Object.
        static void EndProgramm()
        {
            do
            {
                Thread.Sleep(10);

                if (Console.KeyAvailable)
                {
                    chr = Console.ReadKey(true);
                }
            } while (chr.Key != ConsoleKey.C);
        }
    }
}