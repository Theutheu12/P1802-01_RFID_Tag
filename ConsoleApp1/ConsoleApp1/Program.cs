using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace ConsoleApp1
{
    class Program
    {
        //----------------------------------------------------------------------------------------------
        // Constant
        //----------------------------------------------------------------------------------------------
        const string EXCEL_PROG_ID = "Excel.Application";
        const uint MK_E_UNAVAILABLE = 0x800401e3;
        const uint DV_E_FORMATETC = 0x80040064;
        const uint bufferSize = 150;
        const uint scanTimeSize = 20;
        const uint excelTagIDIndex = 2;
        const uint excelScanTimeIndex = 3;

        //----------------------------------------------------------------------------------------------
        // Variables
        //----------------------------------------------------------------------------------------------
        static ConsoleKeyInfo chrCmd;

        //----------------------------------------------------------------------------------------------
        // Class
        //----------------------------------------------------------------------------------------------
        public class TagIDBuffer
        {
            public int bufferIndex;
            public DateTime[,] scanTime = new DateTime[bufferSize, scanTimeSize];
            public int[] tagID = new int[bufferSize];
            public int[] tagCount = new int[bufferSize];
        }

        static void Main(string[] args)
        {
            string date;
            int ID;
            int searchIndex;
            int compareResult;
            int excelIndex = 1;


            DateTime scanTime = new DateTime(1);
            TimeSpan diffTime = new TimeSpan();
            TimeSpan interval = new TimeSpan(0, 0, 15);

            TagIDBuffer tagIDBuffer = new TagIDBuffer
            {
                bufferIndex = 0
            };

            Thread Thr_EndProgramm = new Thread(EndProgramm);
            Thr_EndProgramm.Start();

            Console.WriteLine("**********************************");
            Console.WriteLine("***** VT UHF Tag Reader v0.1 *****");
            Console.WriteLine("**********************************");

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

            //if (null == excelApp)
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

            do
            {
                ID = RandomInt(150);
                scanTime = DateTime.Now;
                date = string.Format("{0:HH:mm:ss.ff}", scanTime);

                Console.Write("Tag scanned : ");
                Console.Write(date);
                Console.Write(" ");
                Console.WriteLine(ID.ToString());

                searchIndex = Array.IndexOf(tagIDBuffer.tagID, ID);
                //Console.WriteLine(searchIndex.ToString());

                if (searchIndex == -1)
                {
                    tagIDBuffer.tagID[tagIDBuffer.bufferIndex] = ID;
                    cell = sheet.Cells[excelIndex, excelTagIDIndex];
                    cell.Value = ID;
                    tagIDBuffer.scanTime[tagIDBuffer.bufferIndex, tagIDBuffer.tagCount[tagIDBuffer.bufferIndex]] = scanTime;
                    cell = sheet.Cells[excelIndex, excelScanTimeIndex];
                    cell.Value = string.Format("{0:HH:mm:ss.ff}", scanTime);
                    tagIDBuffer.tagCount[tagIDBuffer.bufferIndex]++;
                    tagIDBuffer.bufferIndex++;
                    excelIndex++;
                    excelApp.SaveAs();
                }
                else if (searchIndex >= 0)
                {
                    //Console.WriteLine("new count !");
                    diffTime = scanTime.Subtract(tagIDBuffer.scanTime[searchIndex, tagIDBuffer.tagCount[tagIDBuffer.bufferIndex]]);
                    //Console.WriteLine(diffTime.ToString(@"mm\:ss\.ff"));
                    compareResult = TimeSpan.Compare(interval, diffTime);
                    /*Console.WriteLine("{0} {1} {2} (Compare returns {3})",
                                      interval,
                                      compareResult == 1 ? ">" : compareResult == 0 ? "=" : "<",
                                      diffTime, compareResult);*/

                    if (compareResult == -1 | compareResult== 0)
                    {
                        tagIDBuffer.scanTime[searchIndex, tagIDBuffer.tagCount[searchIndex]] = scanTime;
                        cell = sheet.Cells[excelIndex, excelTagIDIndex];
                        cell.Value = ID;
                        cell = sheet.Cells[excelIndex, excelScanTimeIndex];
                        cell.Value = string.Format("{0:HH:mm:ss.ff}", scanTime);
                        tagIDBuffer.tagCount[searchIndex]++;
                        excelIndex++;
                        excelApp.SaveAs();
                    }

                }
                else
                {
                    Console.WriteLine("No case !!!");
                }

                //Console.ReadKey();
                Thread.Sleep(250);

            } while (chrCmd.Key != ConsoleKey.C);

            workbook = excelApp.ActiveWorkbook ?? excelApp.Workbooks.Add();

            for (int i = 0; i < tagIDBuffer.bufferIndex; i++)
            {
                Console.WriteLine("Tag ID : \t" + tagIDBuffer.tagID[i].ToString() + "\t"
                                + tagIDBuffer.tagCount[i] + "\t"
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 0]) + " "
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 1]) + " "
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 2]) + " "
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 3]) + " "
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 4]) + " "
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 5]) + " "
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 6]) + " "
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 7]) + " "
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 8]) + " "
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 9]) + " "
                                /*+ string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 10]) + "\t"
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 11]) + "\t"
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 12]) + "\t"
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 13]) + "\t"
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 14]) + "\t"
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 15]) + "\t"
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 16]) + "\t"
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 17]) + "\t"
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 18]) + "\t"
                                + string.Format("{0:HH:mm:ss.ff}", tagIDBuffer.scanTime[i, 19])*/
                                );
            }

            Console.ReadKey();
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

        static int RandomInt(int randomLimit)
        {
            int randomInt = 1;
            Random random = new Random();
            randomInt = random.Next(1,randomLimit);
            return randomInt;
        }
    }
}
