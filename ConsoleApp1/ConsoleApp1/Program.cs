using System;
using System.Threading;

namespace ConsoleApp1
{
    class Program
    {

        //const string EXCEL_PROG_ID = "Excel.Application";
        //const uint MK_E_UNAVAILABLE = 0x800401e3;
        //const uint DV_E_FORMATETC = 0x80040064;

        static ConsoleKeyInfo chrCmd;

        public class tagIDBuffer
        {

        }



        static void Main(string[] args)
        {

            string date;
            int ID;

            Thread Thr_EndProgramm = new Thread(EndProgramm);
            Thr_EndProgramm.Start();

            Console.WriteLine("**********************************");
            Console.WriteLine("***** VT UHF Tag Reader v0.1 *****");
            Console.WriteLine("**********************************");

            do
            {

                date = getDate();
                ID = randomInt(25);

                Console.WriteLine(date);
                Console.WriteLine(ID.ToString());

                //Console.ReadKey();
                Thread.Sleep(1000);

            } while (chrCmd.Key != ConsoleKey.C);



            /*
            dynamic excelApp = null;
            try
            {
                excelApp = Marshal.GetActiveObject(EXCEL_PROG_ID);
            }
            catch (COMException ex)
            {
                switch ((uint)ex.ErrorCode)
                {
                    case MK_E_UNAVAILABLE:
                    case DV_E_FORMATETC:
                        // Excel n'est pas lancé.
                        Console.WriteLine(ex.ErrorCode);
                        break;

                    default:
                        throw;
                }
            }

            if (null == excelApp)
                excelApp = Activator.CreateInstance(Type.GetTypeFromProgID(EXCEL_PROG_ID));

            if (null == excelApp)
            {
                Console.Write("Unable to start Excel");
                return;
            }

            excelApp.Visible = true;

            dynamic workbook = excelApp.ActiveWorkbook ?? excelApp.Workbooks.Add();
            dynamic sheet = workbook.ActiveSheet;
            dynamic cell = sheet.Cells[1, 1];
            cell.Value = date;*/

        }

        //----------------------------------------------------------------------------------------------
        // Threads
        //----------------------------------------------------------------------------------------------

        //La méthode prend en paramètre un et un seul paramètre de type Object.
        static void EndProgramm()
        {
            do
            {
                Thread.Sleep(10);

                if (Console.KeyAvailable)
                {
                    chrCmd = Console.ReadKey(true);
                }
            } while (chrCmd.Key != ConsoleKey.C);
        }

        //----------------------------------------------------------------------------------------------
        // Methods
        //----------------------------------------------------------------------------------------------

        static string getDate()
        {
            string date = string.Format("{0:HH:mm:ss.ff}", DateTime.Now);
            return date;
        }

        static int randomInt(int randomLimit)
        {
            int randomInt = 1;
            Random random = new Random();
            randomInt = random.Next(randomLimit);
            return randomInt;
        }
    }
}
