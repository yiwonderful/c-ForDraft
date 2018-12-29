using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Net;
using System.Net.Sockets;
using System.Collections;
using WZ.Base.DAL_SQL;
using BS.Alarm.WebApi.Models;
using NPOI.HSSF;
using System.Web;
using System.Web.UI;
using Aspose.Words;

namespace ConsoleApplication1
{
    class Program : BaseRepository_SQL<WarningMsgHead>
    {
        static void Main(string[] args)
        {
            #region
            string dt = string.Empty;
            ////dt =DateTime.Now.ToString("yyyy-MM-dd");
            ////dt = DateTime.Now.ToString("yyyy-MM") +"-"+"01";
            //int year = int.Parse(DateTime.Now.ToString("yyyy"));
            ////dt = DateTime.Now.ToString("MM");
            ////int d = int.Parse(dt);
            //int d = year % 4;
            //int month = int.Parse(DateTime.Now.ToString("MM"));
            //string dd=DateTime.Now.ToString("-30");
            //double a = Math.Round(125.0,2);
            //dt = "%&MD_DepName&连铸炼钢";
            //int a=dt.LastIndexOf('&');
            //dt=dt.Substring(a+1);
           // string[] a = { "2", "3", "4" };
           // StringBuilder sb = new StringBuilder();
           // for(int i=0;i<a.Length;i++){
           //     sb.Append("'" + a[i] + "',");
           // }
           //String strn=sb.ToString().TrimEnd(',');
            //20180621
            //string strn = Environment.CurrentDirectory;
            //if (!Directory.Exists(Environment.CurrentDirectory + "//Test"))
            //{
            //    Directory.CreateDirectory(Environment.CurrentDirectory + "//Test");
            //}
            //dt = strn + "//Test//test" + DateTime.Now.ToString("yyyyMMdd") + "-LL.log";
            //Console.WriteLine(strn);
            //StreamWriter s = null;
            //using (FileStream fs = new FileStream(dt, FileMode.Append, FileAccess.Write, FileShare.Write))
            //{
            //    s = new StreamWriter(fs);
            //    s.WriteLine("12344");
            //    s.Close();
            //}
            //
            //WriteLog("123");
            //string arr = "1,2,3,4";
            //string[] ar = arr.Split(',');
            //for (int i = 0;i < ar.Length; i++ )
            //{
            //    Console.Write((ar[i]+" "));
            //}
            //Guid id =Guid.NewGuid();
            //string s="2018-07";
            //s = s.Replace("-","");
            //object o = new object();
            //o={double[0]};
            double a;
            //float.Parse(a);
            //string date_start = DateTime.Now.Date.ToString();
            //string day = DateTime.Now.Day.ToString();
            //string date_end = DateTime.Now.Date.ToString().Replace(day,"01");
            //Console.WriteLine(date_start);
            //Console.WriteLine(date_end);
            //string aa = DateTime.Now.ToString("2018-05-20 HH:mm:ss");
            //string aaa = "2018-05-20";
            //var list = aaa.Split('-');
            //Console.WriteLine(list[0]);
            //string time = "2018-07";
            //string endTime = time + DateTime.Now.ToString("-dd");
            //string[] month_31 = { "01", "03", "05", "07", "08", "10", "12" };
            //bool ToF = false;
            //string m = DateTime.Now.ToString("MM");
            //if (month_31.Contains(m))
            //{
            //    ToF = true;
            //}

            //string ddd = DealTime("15", "04", "2018");
            string dddd = "2018-07 ";
            string d5 = dddd.Substring(0, 4);
            DateTime d6 = Convert.ToDateTime(dddd);
            //DateTime d6;
            //DateTimeFormatInfo dtFormat = new System.Globalization.DateTimeFormatInfo();

            //dtFormat.ShortDatePattern = "yyyy-MM-dd";

            //d6 = Convert.ToDateTime(dddd, dtFormat);
            string i = "01";
            int tbToStore = int.Parse(i)-1;
            string d7 = DateTime.Now.ToString("-MM-dd");
            string tm = "2018/07/07";
            string tm_1 = Convert.ToDateTime(tm).AddDays(-7).ToString("yyyy-MM-dd");
            object nn = null;
            double a1 = Convert.ToDouble(nn);
            string aaa ="'2018070718544124940733c937a4a51','201807091135204255400ef486d439c','201807091136145086334b93b2f63ae','201807091134075553721696c1044f0'";
            string []a5 = aaa.Split(',');
            string [] a6={"1","2","3","4","5"};
            int double2 = 0;
            string e = "3.0201658913370196e-38";
            //e = "12e-2";
            e = "7e-15";
            double? da=Convert.ToDouble(1132368.1221);
            double? db=Convert.ToDouble(1184619.875);
            //52251.7529
            //double? dc = (double?)(Convert.ToDecimal(1.012) - Convert.ToDecimal(1.0002));
            double? dc = 1.012 - 1.0002;
            //double dc = 1184619.875 - 1132368.1221;
            //double ass = 66.24;
            //double sadf = (double)(Convert.ToDecimal(ass) * Convert.ToDecimal(100.0d));
            //Console.WriteLine("{0:G50}", sadf);
            double ass = 66.24;
            double sadf = ass * 100.0d;
            Console.WriteLine("{0:G50}", sadf);
            Console.WriteLine(dc);

            foreach (var ii in a6)
            {
                if (int.Parse(ii) % 2 == 0)
                {
                    ++double2;
                }
            }
            DateTime time_test = Convert.ToDateTime("2018-04-09 23:05:00.000");
            string time_out = time_test.ToString("yyyy-MM-dd HH:00:ss gg");//24小时
            time_out = time_test.Minute.ToString("00");
            Console.WriteLine(time_out);

            //2018.09.04
            Bday obj = new Bday();
            List<Bday> list = new List<Bday>();
            list.Add(new Bday()
            {
                num = 1,
                num1 = 11,
                id = "01"
            });
            list.Add(new Bday()
            {
                num = 2,
                num1 = 22,
                id = "02"
            });
            double  dou = list.Where(x => x.id == "01").Select(x => x.num1).Sum();
            Console.WriteLine(dou);

            string strr = "12333-3";
            int numm = strr.IndexOf('-');
            #endregion
            //2018.09.18
            decimal fl = 1.00M;
            //fl = (float)Math.Round(fl, 2);
            //float dd = (float)(Math.Round(1 / (float)1 * 100) / 100.00);
            string pattern="^(-){0,1}\\d+(\\.\\d{0,2})?$";
            Regex reg=new Regex(pattern);
            Match ma = reg.Match("-1.25");
            bool bo = ma.Success;
            Console.WriteLine(fl);
            //functionA funa = new functionA();
            //funa.service();
            #region 数据库连接
            string sql = string.Format(@"
            insert into [WarningMsgHead]([ID],[Receiver]) values('03','003'); 
            ");
            int nu = SqlDbHelper.ExecuteNonQuery(sql);
            //BaseRepository_SQL<WarningMsgHead> sl = new BaseRepository_SQL<WarningMsgHead>();
            //int numm1 = sl.SqlExecute(sql, null);
            Console.WriteLine(nu);
            #endregion
            //Console.WriteLine(numm1);
            double ts;
            ts = (double)Math.Round(0.20000000298023224, 4);
            Console.WriteLine(ts);

            //string date = "2018-2";
            //DateTime days = DateTime.Parse(date);
            //int dayNum = System.Threading.Thread.CurrentThread.CurrentUICulture.Calendar.GetDaysInMonth(days.Year, days.Month);
            //Console.WriteLine(dayNum);
            //string numss = Guid.NewGuid().ToString().Replace("-","");
            //Console.WriteLine(numss);
            int team = 0;
            int teamNum = 0;
            team = team == teamNum ? 0 : 5;
            Console.WriteLine(team);
            string date = "2018-11-01";
            int dayN=30;
            int endDate = int.Parse((date.Substring(0, 7) + dayN.ToString()).Replace("-",""));
            DateTime dateLast = DateTime.Parse(date).AddDays(-1);
            int dat = int.Parse( dateLast.ToString("yyyyMMdd"));
            DayOfWeek week = (DateTime.Parse(date)).DayOfWeek;
            var weekdays = new string[] { "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六" };
            string weekday=weekdays[(int)week];
            Console.WriteLine(weekday);
            Console.WriteLine(dat);
            string date111 = "20180105";
            //DateTime dtt = DateTime.Parse(date111);
            string date1111 = DateTime.ParseExact(date111, "yyyyMMdd", null, System.Globalization.DateTimeStyles.AllowWhiteSpaces).ToString("yyyy-MM-dd");
            Console.WriteLine(date1111);
            string datt = DateTime.ParseExact(date111, "yyyyMMdd", null, System.Globalization.DateTimeStyles.AllowWhiteSpaces).AddMonths(-1).ToString("yyyyMM");
            Console.WriteLine(datt);
            #region Export_Word
            ExportToWorld exp = new ExportToWorld();
            exp.ExportWord();
            exp.MergeCell();
            #endregion
            #region 2018-12-30
            Console.WriteLine("2018-12-30:");
            int testRef = 2;
            refTest(ref testRef);
            Console.WriteLine(testRef);
            Func<int,string> RetBook = new Func<int,string>(FuncBook);
            Console.WriteLine(RetBook(1));
            Func<string> funcValue = delegate
            {
                return "我是即将传递的值3";
            };
            DisPlayValue(funcValue);
            int temp = int.MaxValue;
            try
            {
                //int a_t = unchecked(temp * 2);
                int a_t = checked(temp * 2);
                Console.WriteLine(a_t);
            }
            catch (OverflowException)
            {
                Console.WriteLine("溢出了，要处理哟");
            }
            #endregion
            Console.ReadLine();
        }

        private static void DisPlayValue(Func<string> func)
        {
            string RetFunc = func();
            Console.WriteLine("我在测试一下传过来值：{0}", RetFunc);
        }

        public static string FuncBook(int name)
        {
            return "送书来了，书数量为" + name + "本";
        }
        public static void refTest( ref int i)
        {
            i = 123;
        }
        public static void testSql()
        {
            //int num=SqlExecute()
        }

        public static string DealTime(string day_this, string month, string year)
        {
            string dateTime = string.Empty;
            string[] month_31 = { "01", "03", "05", "07", "08", "10", "12" };
            string[] month_30 = { "04", "06", "09", "11" };
            if (month_31.Contains(month))
            {
                dateTime = year + "-" + month + "-" + day_this;
            }
            else if (month_30.Contains(month))
            {
                if (day_this == "31")
                    dateTime = year + "-" + month + "-" + "30";
                else
                    dateTime = year + "-" + month + "-" + day_this;
            }
            else
            {
                if (int.Parse(year) % 4 == 0)
                {
                    if (int.Parse(day_this) > 29)
                        dateTime = year + "-" + month + "-29";
                    else
                        dateTime = year + "-" + month + "-" + day_this;
                }
                else
                {
                    if (int.Parse(day_this) > 28)
                        dateTime = year + "-" + month + "-28";
                    else
                        dateTime = year + "-" + month + "-" + day_this;
                }
            }
            return dateTime;
        }

        public static void WriteLog(string log)
        {
            if (!Directory.Exists(Environment.CurrentDirectory + "//Logs"))
            {
                Directory.CreateDirectory(Environment.CurrentDirectory + "//Logs");
            }
            string filePath = Environment.CurrentDirectory + "//Logs//Log" + DateTime.Now.ToString("yyyyMMdd") + "-WCF.log";
            try
            {
                StreamWriter sw = null;
                using (FileStream fs = new FileStream(filePath, FileMode.Append, FileAccess.Write, FileShare.Write))
                {
                    sw = new StreamWriter(fs);
                    sw.WriteLine("[" + DateTime.Now.ToString("HH:mm:ss.fff") + "] " + log);
                    sw.Close();
                }
            }
            catch
            {
            }
        }


        public static Decimal ChangeToDecimal(string strData)
        {
            Decimal dData = 0.0M;
            if (strData.Contains("E") || strData.Contains("e"))
            {
                dData = Convert.ToDecimal(Decimal.Parse(strData.ToString(), System.Globalization.NumberStyles.Float));
            }
            else
            {
                dData = Convert.ToDecimal(strData);
            }
            //if (dData == 0)
            //{
            //    dData = (Decimal)("");
            //}
            return Math.Round(dData, 4);
        }
    }
    public  class Bday
    {
        public double num{get;set;}

        public double num1 { get; set; }

        public string id { get; set; }
    }


    public class functionA
    {
        public void service()
        {
            int port = 6000;
            string host = "192.168.0.35";

            IPAddress ip = IPAddress.Parse(host);
            IPEndPoint ipe = new IPEndPoint(ip, port);

            Socket sSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            sSocket.Bind(ipe);
            sSocket.Listen(10);
            Console.WriteLine("监听已经打开，请等待");

            //receive message
            Socket serverSocket = sSocket.Accept();
            Console.WriteLine("连接已经建立");
            string recStr = "";
            byte[] recByte = new byte[4096];
            int bytes = serverSocket.Receive(recByte, recByte.Length, 0);
            recStr += Encoding.ASCII.GetString(recByte, 0, bytes);

            //send message
            Console.WriteLine("服务器端获得信息:{0}", recStr);
            string sendStr = "send to client :hello";
            byte[] sendByte = Encoding.ASCII.GetBytes(sendStr);
            serverSocket.Send(sendByte, sendByte.Length, 0);
            //serverSocket.Close();
            //sSocket.Close();
        }
            
    }

    public class functionB
    {
        Socket clientSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
        public void client()
        {
            int port = 6000;
            string host = "127.0.0.1";//服务器端ip地址

            IPAddress ip = IPAddress.Parse(host);
            IPEndPoint ipe = new IPEndPoint(ip, port);

            //Socket clientSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            clientSocket.Connect(ipe);

            //send message
            string sendStr = "send to server : hello,ni hao";
            byte[] sendBytes = Encoding.ASCII.GetBytes(sendStr);
            clientSocket.Send(sendBytes);

            //receive message
            string recStr = "";
            byte[] recBytes = new byte[4096];
            int bytes = clientSocket.Receive(recBytes, recBytes.Length, 0);
            recStr += Encoding.ASCII.GetString(recBytes, 0, bytes);
            Console.WriteLine(recStr);

            clientSocket.Close();
        }

        public void Connect(IPAddress ip, int port)
        {
            this.clientSocket.BeginConnect(ip, port, new AsyncCallback(ConnectCallback), this.clientSocket);
        }

        private void ConnectCallback(IAsyncResult ar)
        {
            try
            {
                Socket handler = (Socket)ar.AsyncState;
                handler.EndConnect(ar);
            }
            catch (SocketException ex)
            { }
        }

        #region 发送的方法
        public void Send(string data)
        {
            Send(System.Text.Encoding.UTF8.GetBytes(data));
        }

        private void Send(byte[] byteData)
        {
            try
            {
                int length = byteData.Length;
                byte[] head = BitConverter.GetBytes(length);
                byte[] data = new byte[head.Length + byteData.Length];
                Array.Copy(head, data, head.Length);
                Array.Copy(byteData, 0, data, head.Length, byteData.Length);
                this.clientSocket.BeginSend(data, 0, data.Length, 0, new AsyncCallback(SendCallback), this.clientSocket);
            }
            catch (SocketException ex)
            { }
        }

        private void SendCallback(IAsyncResult ar)
        {
            try
            {
                Socket handler = (Socket)ar.AsyncState;
                handler.EndSend(ar);
            }
            catch (SocketException ex)
            { }
        }
        #endregion

        #region 接收的方法
        byte[] MsgBuffer = new byte[1024];
        public void ReceiveData()
        {
            clientSocket.BeginReceive(MsgBuffer, 0, MsgBuffer.Length, 0, new AsyncCallback(ReceiveCallback), null);
        }

        private void ReceiveCallback(IAsyncResult ar)
        {
            try
            {
                int REnd = clientSocket.EndReceive(ar);
                if (REnd > 0)
                {
                    byte[] data = new byte[REnd];
                    Array.Copy(MsgBuffer, 0, data, 0, REnd);

                    //在此次可以对data进行按需处理

                    clientSocket.BeginReceive(MsgBuffer, 0, MsgBuffer.Length, 0, new AsyncCallback(ReceiveCallback), null);
                }
                else
                {
                    dispose();
                }
            }
            catch (SocketException ex)
            { }
        }

        private void dispose()
        {
            try
            {
                this.clientSocket.Shutdown(SocketShutdown.Both);
                this.clientSocket.Close();
            }
            catch (Exception ex)
            { }
        }
        #endregion

        #region 处理接收的包
        public Hashtable DataTable = new Hashtable();//因为要接收到多个服务器（ip）发送的数据，此处按照ip地址分开存储发送数据

        public void DataArrial(byte[] Data, string ip)
        {
            try
            {
                if (Data.Length < 12)//按照需求进行判断
                {
                    lock (DataTable)
                    {
                        if (DataTable.Contains(ip))
                        {
                            DataTable[ip] = Data;
                            return;
                        }
                    }
                }

                if (Data[0] != 0x1F || Data[1] != 0xF1)//标志位（按照需求编写）
                {
                    if (DataTable.Contains(ip))
                    {
                        if (DataTable != null)
                        {
                            byte[] oldData = (byte[])DataTable[ip];//取出粘包数据
                            if (oldData[0] != 0x1F || oldData[1] != 0xF1)
                            {
                                return;
                            }
                            byte[] newData = new byte[Data.Length + oldData.Length];
                            Array.Copy(oldData, 0, newData, 0, oldData.Length);
                            Array.Copy(Data, 0, newData, oldData.Length, Data.Length);//组成新数据数组，先到的数据排在前面，后到的数据放在后面

                            lock (DataTable)
                            {
                                DataTable[ip] = null;
                            }
                            DataArrial(newData, ip);
                            return;
                        }
                    }
                    return;
                }

                int revDataLength = Data[2];//打算发送数据的长度
                int revCount = Data.Length;//接收的数据长度
                if (revCount > revDataLength)//如果接收的数据长度大于发送的数据长度，说明存在多帧数据，继续处理
                {
                    byte[] otherData = new byte[revCount - revDataLength];
                    Data.CopyTo(otherData, revCount - 1);
                    Array.Copy(Data, revDataLength, otherData, 0, otherData.Length);
                    Data = (byte[])Redim(Data, revDataLength);
                    DataArrial(otherData, ip);
                }
                if (revCount < revDataLength) //接收到的数据小于要发送的长度
                {
                    if (DataTable.Contains(ip))
                    {
                        DataTable[ip] = Data;//更新当前粘包数据
                        return;
                    }
                }

                //此处可以按需进行数据处理
            }
            catch (Exception ex)
            { }
        }

        private Array Redim(Array origArray, Int32 desizedSize)
        {
            //确认每个元素的类型
            Type t = origArray.GetType().GetElementType();
            //创建一个含有期望元素个数的新数组
            //新数组的类型必须匹配数组的类型
            Array newArray = Array.CreateInstance(t, desizedSize);
            //将原数组中的元素拷贝到新数组中。
            Array.Copy(origArray, 0, newArray, 0, Math.Min(origArray.Length, desizedSize));
            //返回新数组
            return newArray;
        }
        #endregion 
    }

    public class ExportToWorld
    {
        public void ExportWord()
        {
            string filePath = System.AppDomain.CurrentDomain.BaseDirectory + "/Template.doc"; ;
            string filePath1 = System.AppDomain.CurrentDomain.BaseDirectory + "/Template1.doc"; ;
            //预先生成数据
            List<Student> studentData = new List<Student>();
            for (int i = 0; i < 100; i++)
            {
                studentData.Add(new Student()
                {
                    name = "学生" + i.ToString("D3"),
                    schoolName = "某某中学",
                    num = i,
                    score = i
                });
            }
            //加载word模板。
            Aspose.Words.Document doc = new Aspose.Words.Document(filePath);
            Aspose.Words.DocumentBuilder docWriter = new Aspose.Words.DocumentBuilder(doc);

            double[] colWidth = new double[] { 45, 60, 33, 55 };
            string[] colName = new string[] { "编号", "姓名", "分数", "学校" };
            int pageSize = 0;
            for (int i = 0, j = studentData.Count; i < j; i++)
            {
                if (pageSize == 0)
                {
                    //word页刚开始，一个表格的开始，要插入一个表头
                    docWriter.InsertBreak(Aspose.Words.BreakType.ParagraphBreak);
                    docWriter.StartTable();
                    AsposeCreateCell(docWriter, colWidth[0], colName[0]);
                    AsposeCreateCell(docWriter, colWidth[1], colName[1]);
                    AsposeCreateCell(docWriter, colWidth[2], colName[2]);
                    AsposeCreateCell(docWriter, colWidth[3], colName[3]);
                    docWriter.EndRow();
                }
                else if (pageSize == 30)//经过测算，每页word中可以放置30行
                {
                    //结束第一个表格，插入分栏符号，并开始另一个表格
                    docWriter.EndTable();
                    docWriter.InsertBreak(Aspose.Words.BreakType.ColumnBreak);
                    docWriter.InsertBreak(Aspose.Words.BreakType.ParagraphBreak);
                    docWriter.StartTable();
                    AsposeCreateCell(docWriter, colWidth[0], colName[0]);
                    AsposeCreateCell(docWriter, colWidth[1], colName[1]);
                    AsposeCreateCell(docWriter, colWidth[2], colName[2]);
                    AsposeCreateCell(docWriter, colWidth[3], colName[3]);
                    docWriter.EndRow();
                }
                else if (pageSize == 60)//word分栏为2栏，那么一页word可以放60行数据
                {
                    //一页word完毕，关闭表格，并绘制下一页的表头。
                    docWriter.EndTable();
                    docWriter.InsertBreak(Aspose.Words.BreakType.PageBreak);
                    docWriter.InsertBreak(Aspose.Words.BreakType.ParagraphBreak);
                    docWriter.StartTable();
                    AsposeCreateCell(docWriter, colWidth[0], colName[0]);
                    AsposeCreateCell(docWriter, colWidth[1], colName[1]);
                    AsposeCreateCell(docWriter, colWidth[2], colName[2]);
                    AsposeCreateCell(docWriter, colWidth[3], colName[3]);
                    docWriter.EndRow();
                    pageSize = 0;
                }
                pageSize++;
                //创建内容
                AsposeCreateCell(docWriter, colWidth[0], studentData[i].num.ToString());
                AsposeCreateCell(docWriter, colWidth[1], studentData[i].name);
                AsposeCreateCell(docWriter, colWidth[2], studentData[i].score.ToString());
                AsposeCreateCell(docWriter, colWidth[3], studentData[i].schoolName);
                docWriter.EndRow();
            }//end for
            //保存文件
            doc.Save(filePath1, Aspose.Words.SaveFormat.Doc);
        }

        public void AsposeCreateCell(Aspose.Words.DocumentBuilder builder, double width, string text)
        {

            builder.InsertCell();
            builder.CellFormat.Borders.LineStyle = Aspose.Words.LineStyle.Single;
            builder.CellFormat.Borders.Color = System.Drawing.Color.Black;
            builder.CellFormat.Width = width;//单元格的宽度
            builder.CellFormat.LeftPadding = 3;//单元格的左内边距
            builder.CellFormat.RightPadding = 3;//单元格的右内边距
            builder.RowFormat.Height = 20;//行高
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐
            builder.ParagraphFormat.Alignment = Aspose.Words.ParagraphAlignment.Center;//水平居中对齐
            builder.Write(text);
        }

        public void MergeCell()
        {
            string filePath = System.AppDomain.CurrentDomain.BaseDirectory+"/Template.doc";
            string filePath1 = System.AppDomain.CurrentDomain.BaseDirectory + "/Template1.doc"; ;
            Aspose.Words.Document doc = new Aspose.Words.Document(filePath);
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
            builder.InsertCell();
            builder.CellFormat.Borders.LineStyle = LineStyle.Single;
            builder.CellFormat.Borders.Color = System.Drawing.Color.Black;
            //水平合并
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.First;
            //垂直合并
            //builder.CellFormat.HorizontalMerge= Aspose.Words.Tables.CellMerge.First;
            builder.Write("Text in merged cells.");

            builder.InsertCell();
            builder.CellFormat.Borders.LineStyle = LineStyle.Single;
            builder.CellFormat.Borders.Color = System.Drawing.Color.Black;
            //builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.Write("Text in one cell");
            builder.EndRow();

            builder.InsertCell();
            builder.CellFormat.Borders.LineStyle = LineStyle.Single;
            builder.CellFormat.Borders.Color = System.Drawing.Color.Black;
            // 此单元格垂直合并到单元格上方，并应为空.
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.Previous;
            //builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.None;

            builder.InsertCell();
            builder.CellFormat.Borders.LineStyle = LineStyle.Single;
            builder.CellFormat.Borders.Color = System.Drawing.Color.Black;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.Write("Text in another cell");
            builder.EndRow();
            doc.Save(filePath1, Aspose.Words.SaveFormat.Doc);
        }

        public class Student{
            public string name { get; set; }

            public string schoolName { get; set; }

            public int num { get; set; }

            public decimal score { get; set; }
        }
    }
    public class ParallelInvoke
    {
        /// <summary>
        /// Invoke方式一 action
        /// </summary>
        public void Client1()
        {
            //Stopwatch 可以测量一个时间间隔的运行时间，也可以测量多个时间间隔的总运行时间。一般用来测量代码执行所用的时间或者计算性能数据，在优化代码性能上可以使用Stopwatch来测量时间。
            Stopwatch stopWatch = new Stopwatch();

            Console.WriteLine("主线程：{0}线程ID ： {1}；开始", "Client1", Thread.CurrentThread.ManagedThreadId);
            stopWatch.Start();
            Parallel.Invoke(() => Task1("task1"), () => Task2("task2"), () => Task3("task3"));
            stopWatch.Stop();
            Console.WriteLine("主线程：{0}线程ID ： {1}；结束,共用时{2}ms", "Client1", Thread.CurrentThread.ManagedThreadId, stopWatch.ElapsedMilliseconds);
        }

        private void Task1(string data)
        {
            Thread.Sleep(5000);
            Console.WriteLine("任务名：{0}线程ID ： {1}", data, Thread.CurrentThread.ManagedThreadId);
        }

        private void Task2(string data)
        {
            Console.WriteLine("任务名：{0}线程ID ： {1}", data, Thread.CurrentThread.ManagedThreadId);
        }

        private void Task3(string data)
        {
            Console.WriteLine("任务名：{0}线程ID ： {1}", data, Thread.CurrentThread.ManagedThreadId);
        }
    }
}
