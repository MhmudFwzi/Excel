using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net;
using System.Net.Sockets;
using ExcelLibrary.SpreadSheet;
using ExcelLibrary.CompoundDocumentFormat;

namespace CompleteTrial
{
    class Program
    {
        private TcpListener tcpListen;
        public static void Main(string[] args)
        {

            Program program = new Program();
            program.StartServer();

            while (true) ;

        }

        private bool StartServer()
        {
            //IPAddress ipAddress = Dns.GetHostEntry("localhost").AddressList[0];


            try
            {
                // Creating new TCP Listener on Port 4532
                tcpListen = new TcpListener(IPAddress.Any, 4532);
                tcpListen.Start();
                tcpListen.BeginAcceptTcpClient(new AsyncCallback(this.ProcessEvents), tcpListen);
                //tcpListen.BeginAcceptSocket(new AsyncCallback(this.ProcessEvents), tcpListen);

                Console.WriteLine("Listing at Port {0}.", 2020);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return false;
            }

            return true;
        }

        private void ProcessEvents(IAsyncResult asyn)
        {
            try
            {
                TcpListener processListen = (TcpListener)asyn.AsyncState;
                TcpClient tcpClient = processListen.EndAcceptTcpClient(asyn);
                NetworkStream myStream = tcpClient.GetStream();
                if (myStream.CanRead)
                {
                    StreamReader readerStream = new StreamReader(myStream); //Creating new StreamReader
                    string myMessage = readerStream.ReadToEnd();    //Reading the stream till the end
                    Console.WriteLine(myMessage);
                    //string prs = Console.ReadLine();
                    string path = "Ordersfile.xls";
                   if (!(File.Exists(path)))
                    {
                        
                        Workbook iOrdersFile = new Workbook();
                        Worksheet iordersSheet = new Worksheet("ordersSheet");
                        iOrdersFile.Worksheets.Add(iordersSheet);
                        iordersSheet.Cells[0, 0] = new Cell("order no.");
                        iordersSheet.Cells[0, 1] = new Cell("station no.");
                        iordersSheet.Cells[0, 2] = new Cell("product no.");
                        iordersSheet.Cells[0, 3] = new Cell("Date&Time");
                        iOrdersFile.Save("Ordersfile.xls");
                        
                    }
                    int i = 1;
                    Workbook OrdersFile = Workbook.Load("Ordersfile.xls");
                    Worksheet ordersSheet = OrdersFile.Worksheets[0];
                    //Cell indexcheck = new Cell(i.ToString());
                    while (ordersSheet.Cells[i, 0].ToString() == i.ToString())
                    {
                        i++;
                    }
                    for (int j = 0; j < myMessage.Length; j+=5)
                    {
                        ordersSheet.Cells[i, 0] = new Cell(i.ToString());
                        ordersSheet.Cells[i, 1] = new Cell(myMessage.Substring(j, 2));
                        ordersSheet.Cells[i, 2] = new Cell(myMessage.Substring(j+2, 2));
                        ordersSheet.Cells[i, 3] = new Cell(DateTime.Now.ToString());
                        i++;
                    }
                   
                    OrdersFile.Save("Ordersfile.xls");
                    readerStream.Close();
                }
                myStream.Close();
                tcpClient.Close();
                tcpListen.BeginAcceptTcpClient(new AsyncCallback(this.ProcessEvents), tcpListen);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }
    }
}

