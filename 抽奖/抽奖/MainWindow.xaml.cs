using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using System.Data.Odbc;
using System.Data;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Timers;
using System.Threading;
using System.Collections;


namespace 抽奖
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        System.Data.DataTable dt = new System.Data.DataTable();
        ExcelHelper myExcelHelper;
        string[] OneTwoPrize = new string[1];
        string[] ThreeFourFivePrize = new string[1];
        string[] SixSevenPrize = new string[1];

        string[] RandomOneTwoPrize = new string[1];
        string[] RandomThreeFourFivePrize = new string[1];
        string[] RandomSixSevenPrize = new string[1];



        #region 导入数据
        private void ImportNameList_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.InitialDirectory = "d:\\";
            openFileDialog.Filter = "Microsoft Excel files(*.xls)|*.xls;*.xlsx";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            bool? result = openFileDialog.ShowDialog();
            if (result == true)
            {
                dt = new System.Data.DataTable();
                myExcelHelper = new ExcelHelper(openFileDialog.FileName);
                dt = myExcelHelper.ExcelToDataTable("抽奖人员", true);

                if (dt != null)
                {
                    MessageBox.Show("抽奖名单导入成功！");

                    GetNameList();

                    //打乱抽奖顺序
                    RandomThreeFourFivePrize = MakeNameListRandom(ThreeFourFivePrize);
                    RandomSixSevenPrize = MakeNameListRandom(SixSevenPrize);
                    RandomOneTwoPrize = MakeNameListRandomForOneTwoPrize(OneTwoPrize);


                    //导出抽奖名单
                    ExportExcel();

                    NameList.Visibility = Visibility.Collapsed;
                    FirstPage.Visibility = Visibility.Visible;
                    LittleGame.Visibility = Visibility.Visible;
                    LuckyDraw.Visibility = Visibility.Visible;
                    Setting.Visibility = Visibility.Visible;

                    FirstPage.IsSelected = true;


                }
                else
                {
                    MessageBox.Show("导入失败！请重新导入");
                }
            }
        }
        #endregion

        #region 制作抽奖名单
        public void GetNameList()
        {
            int totalPeople = dt.Rows.Count;
            string[] name = new string[totalPeople];
            int[] number = new int[totalPeople];

            //获得人名数组 和 票数数组
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow myDataRow = dt.Rows[i];
                name[i] = myDataRow[0].ToString();
                number[i] = Convert.ToInt32(myDataRow[1]);
            }

            //获得总票数
            int totalTickets = 0;
            for (int i = 0; i < number.Length; i++)
            {
                totalTickets += number[i];
            }

            //获得拥有一票以上的人数
            int AtLeastOneTicketNumber = 0;
            for (int i = 0; i < number.Length; i++)
            {
                if (number[i] >= 1)
                {
                    AtLeastOneTicketNumber++;
                }
            }

            //设置五个抽奖名单的数组大小
            OneTwoPrize = new string[totalTickets];
            ThreeFourFivePrize = new string[AtLeastOneTicketNumber];
            SixSevenPrize = new string[totalPeople];

            RandomOneTwoPrize = new string[totalTickets];
            RandomThreeFourFivePrize = new string[AtLeastOneTicketNumber];
            RandomSixSevenPrize = new string[totalPeople];

            //制作六、七等奖的抽奖名单数组
            for (int i = 0; i < SixSevenPrize.Length; i++)
            {
                SixSevenPrize[i] = name[i];
            }

            //制作三、四、五等奖的抽奖名单数组
            int NotZero = 0;
            for (int i = 0; i < name.Length; i++)
            {
                if (number[i] >= 1)
                {
                    ThreeFourFivePrize[NotZero] = name[i];
                    NotZero++;
                }
            }


            //制作一、二等奖的抽奖名单数组
            int count = 0;
            for (int i = 0; i < name.Length; i++)
            {
                if (number[i] > 0)
                {

                    for (int j = 0; j < number[i]; j++)
                    {
                        OneTwoPrize[count + j] = name[i];
                    }
                    count = count + number[i];
                }
            }

        }
        #endregion

        #region 抽奖名单乱序
        public string[] MakeNameListRandom(string[] stringArray)
        {
            string[] oldStringArray = new string[stringArray.Length];
            for (int i = 0; i < oldStringArray.Length; i++)
            {
                oldStringArray[i] = stringArray[i];
            }

            string[] newStringArray = new string[oldStringArray.Length];
            int k = oldStringArray.Length;

            for (int i = 0; i < oldStringArray.Length; i++)
            {
                int temp = new Random().Next(0, k);

                newStringArray[i] = oldStringArray[temp];

                for (int j = temp; j < oldStringArray.Length - 1; j++)
                {
                    oldStringArray[j] = oldStringArray[j + 1];
                }
                k--;
            }

            return newStringArray;
        }

        
        public string[] MakeNameListRandomForOneTwoPrize(string[] stringArray)
        {
            string[] oldStringArray = new string[stringArray.Length];
            for (int i = 0; i < oldStringArray.Length; i++)
            {
                oldStringArray[i] = stringArray[i];
            }

            string[] newStringArray = new string[oldStringArray.Length];
            int k = oldStringArray.Length;

            for (int i = 0; i < oldStringArray.Length; i++)
            {
                int temp = new Random().Next(0, k);


                if ((i > 2)&& AtLeastFourDifferentNames(oldStringArray, k))
                {
                    if ((oldStringArray[temp] != newStringArray[i - 1])&& (oldStringArray[temp] != newStringArray[i - 2]) && (oldStringArray[temp] != newStringArray[i - 3]))
                    {
                        newStringArray[i] = oldStringArray[temp];

                        for (int j = temp; j < oldStringArray.Length - 1; j++)
                        {
                            oldStringArray[j] = oldStringArray[j + 1];
                        }
                        k--;
                    }
                    else
                    {
                        i--;
                    }
                }
                else if(i==1)
                {
                    if (oldStringArray[temp] != newStringArray[0]) 
                    {
                        newStringArray[i] = oldStringArray[temp];

                        for (int j = temp; j < oldStringArray.Length - 1; j++)
                        {
                            oldStringArray[j] = oldStringArray[j + 1];
                        }
                        k--;
                    }
                    else
                    {
                        i--;
                    }
                }
                else if (i == 2)
                {
                    if ((oldStringArray[temp] != newStringArray[0])&&(oldStringArray[temp] != newStringArray[1]))
                    {
                        newStringArray[i] = oldStringArray[temp];

                        for (int j = temp; j < oldStringArray.Length - 1; j++)
                        {
                            oldStringArray[j] = oldStringArray[j + 1];
                        }
                        k--;
                    }
                    else
                    {
                        i--;
                    }
                }
                else
                {
                    newStringArray[i] = oldStringArray[temp];

                    for (int j = temp; j < oldStringArray.Length - 1; j++)
                    {
                        oldStringArray[j] = oldStringArray[j + 1];
                    }
                    k--;
                }

            }

            return newStringArray;
        }

        public bool AtLeastFourDifferentNames(string[] myStringArray,int length)
        {
            string nameA = "";
            string nameB = "";
            string nameC = "";
            string nameD = "";

            if(length>=3)
            {
                nameA = myStringArray[0];

                for(int i=0;i<length+1;i++)
                {
                    if(myStringArray[i]!=nameA)
                    {
                        nameB = myStringArray[i];
                        break;
                    }
                }

                if(nameB!="")
                {
                    for(int i=0;i<length+1;i++)
                    {
                        if ((myStringArray[i] != nameA)&&(myStringArray[i]!=nameB))
                        {
                            nameC = myStringArray[i];
                            break;
                        }
                    }

                    if(nameC!="")
                    {
                        for(int i=0;i<length+1;i++)
                        {
                            if ((myStringArray[i] != nameA) && (myStringArray[i] != nameB) && (myStringArray[i] != nameC))
                            {
                                nameD = myStringArray[i];
                                break;
                            }
                        }

                        if(nameD!="")
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }

        }
        #endregion

        #region 导出数据（用于测试）
        string str_fileName;                                                  //定义变量Excel文件名
        Microsoft.Office.Interop.Excel.Application ExcelApp;                  //声明Excel应用程序
        Workbook ExcelDoc;                                                    //声明工作簿
        Worksheet ExcelSheet;                                                 //声明工作表

        void ExportExcel()
        {
            //创建excel模板
            str_fileName = "d:\\" + "抽奖名单 " + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";    //文件保存路径及名称
            ExcelApp = new Microsoft.Office.Interop.Excel.Application();                          //创建Excel应用程序 ExcelApp
            ExcelDoc = ExcelApp.Workbooks.Add(Type.Missing);                                      //在应用程序ExcelApp下，创建工作簿ExcelDoc
            ExcelSheet = ExcelDoc.Worksheets.Add(Type.Missing);                                   //在工作簿ExcelDoc下，创建工作表ExcelSheet

            //设置Excel列名           
            ExcelSheet.Cells[1, 1] = "一、二等奖";
            ExcelSheet.Cells[1, 2] = "三、四、五等奖";
            ExcelSheet.Cells[1, 3] = "六、七等奖";
            ExcelSheet.Cells[1, 4] = "一、二等奖随机";
            ExcelSheet.Cells[1, 5] = "三、四、五等奖随机";
            ExcelSheet.Cells[1, 6] = "六、七等奖随机";

            //输出各个参数值
            for (int i = 0; i < OneTwoPrize.Length; i++)
            {
                ExcelSheet.Cells[2 + i, 1] = OneTwoPrize[i].ToString();
                ExcelSheet.Cells[2 + i, 4] = RandomOneTwoPrize[i].ToString();
            }

            for (int i = 0; i < ThreeFourFivePrize.Length; i++)
            {
                ExcelSheet.Cells[2 + i, 2] = ThreeFourFivePrize[i].ToString();
                ExcelSheet.Cells[2 + i, 5] = RandomThreeFourFivePrize[i].ToString();
            }

            for (int i = 0; i < SixSevenPrize.Length; i++)
            {
                ExcelSheet.Cells[2 + i, 3] = SixSevenPrize[i].ToString();
                ExcelSheet.Cells[2 + i, 6] = RandomSixSevenPrize[i].ToString();

            }

            MessageBox.Show("抽奖名单打乱成功！");

            ExcelSheet.SaveAs(str_fileName);                                                      //保存Excel工作表
            ExcelDoc.Close(Type.Missing, str_fileName, Type.Missing);                             //关闭Excel工作簿
            ExcelApp.Quit();                                                                      //退出Excel应用程序    
        }

        void ExportExcelSub(string[] myStringArray)
        {

            //创建excel模板
            str_fileName = "d:\\" + "子抽奖名单 " + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";    //文件保存路径及名称
            ExcelApp = new Microsoft.Office.Interop.Excel.Application();                          //创建Excel应用程序 ExcelApp
            ExcelDoc = ExcelApp.Workbooks.Add(Type.Missing);                                      //在应用程序ExcelApp下，创建工作簿ExcelDoc
            ExcelSheet = ExcelDoc.Worksheets.Add(Type.Missing);                                   //在工作簿ExcelDoc下，创建工作表ExcelSheet

            //设置Excel列名           
            ExcelSheet.Cells[1, 1] = "子抽奖名单";

            //输出各个参数值
            for (int i = 0; i < myStringArray.Length; i++)
            {
                ExcelSheet.Cells[2 + i, 1] = myStringArray[i].ToString();
            }

            MessageBox.Show("成功导出Excel");

            ExcelSheet.SaveAs(str_fileName);                                                      //保存Excel工作表
            ExcelDoc.Close(Type.Missing, str_fileName, Type.Missing);                             //关闭Excel工作簿
            ExcelApp.Quit();                                                                      //退出Excel应用程序    
        }

        #endregion

        #region 小游戏
        System.Timers.Timer littleGameTimer = new System.Timers.Timer(100);
        bool isSix = false;
        int gameCount = 0;
        int nameListNumber = 0;
        string littleGameResult = "";
        int[] numberList = new int[1];
        string[] subRandomSixSevenPrize = new string[1];


        #region 小游戏 开始/停止
        private void StartLittleGame_Click(object sender, RoutedEventArgs e)
        {
            if (StartLittleGame.Content.ToString() == "开始")
            {
                StartLittleGame.Content = "停止";
                LittleGameNameResult.Text = "";
                littleGameResult = "";

                if(nameListNumber!=0)
                {
                    nameListNumber--;
                }

                littleGameTimer.Start();
            }
            else if (StartLittleGame.Content.ToString() == "停止")
            {
                StartLittleGame.Content = "开始";
                littleGameTimer.Stop();

                if (isSix == false)
                {
                    if (gameCount == 0)
                    {
                        LittleGameShowNameList(RandomSixSevenPrize, 12);

                        subRandomSixSevenPrize = new string[RandomSixSevenPrize.Length - 12];
                        subRandomSixSevenPrize = SubStringArray(RandomSixSevenPrize, numberList);

                        gameCount = 1;

                        //ExportExcelSub(subRandomSixSevenPrize); //for test
                    }
                    else if (gameCount == 1)
                    {
                        LittleGameShowNameList(subRandomSixSevenPrize, 12);
                        GetSubStringArray(12);

                        StartLittleGame.IsEnabled = false;
                        NextLittleGame.IsEnabled = true;

                        //ExportExcelSub(subRandomSixSevenPrize); //for test
                    }

                }
                else
                {
                    if (gameCount == 0)
                    {
                        LittleGameShowNameList(subRandomSixSevenPrize, 10);
                        GetSubStringArray(10);

                        gameCount = 1;

                        //ExportExcelSub(subRandomSixSevenPrize); //for test
                    }
                    else if (gameCount == 1)
                    {
                        LittleGameShowNameList(subRandomSixSevenPrize, 10);

                        StartLittleGame.IsEnabled = false;
                        NextLittleGame.IsEnabled = false;
                    }
                }
            }
        }

        public void LittleGameShowNameList(string[] myStringArray,int number)
        {
            numberList = new int[number];
            numberList = GetLuckyNameList(myStringArray, number, nameListNumber);

            for (int i = 0; i < numberList.Length; i++)
            {
                littleGameResult += myStringArray[numberList[i]] + " ";
            }

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                LittleGameNameResult.Text = littleGameResult;

            }));
        }

        public void GetSubStringArray(int number)
        {
            string[] tempStringArray = new string[subRandomSixSevenPrize.Length - number];
            tempStringArray = SubStringArray(subRandomSixSevenPrize, numberList);

            subRandomSixSevenPrize = new string[tempStringArray.Length];
            for (int i = 0; i < subRandomSixSevenPrize.Length; i++)
            {
                subRandomSixSevenPrize[i] = tempStringArray[i];
            }
        }

        private void littleGameTimerHandle(object source, ElapsedEventArgs e)
        {
            if (isSix == false)
            {
                if (gameCount == 0)
                {
                    LittleGameFlow(RandomSixSevenPrize);
                }
                else if (gameCount == 1)
                {
                    LittleGameFlow(subRandomSixSevenPrize);
                }
            }
            else
            {
                LittleGameFlow(subRandomSixSevenPrize);
            }
        }

        public void LittleGameFlow(string[] myStringArray)
        {
            if (nameListNumber >= myStringArray.Length)
            {
                nameListNumber = 0;
            }

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                LittleGameNameFlow.Text = myStringArray[nameListNumber++];

            }));
        }

        public int[] GetLuckyNameList(string[] myStringArray, int myNumber, int myNameListNumber)
        {
            int[] result = new int[myNumber];
            int k = 0;

            for (int i = 0; i < myNumber; i++)
            {
                if (myNameListNumber == 0)
                {
                    result[i] = i;
                }
                else
                {
                    if (myNameListNumber - 1 + i < myStringArray.Length - 1)
                    {
                        result[i] = myNameListNumber - 1 + i;
                    }
                    else if (nameListNumber - 1 + i == myStringArray.Length - 1)
                    {
                        result[i] = myNameListNumber - 1 + i;
                        k = 0;
                    }
                    else
                    {
                        result[i] = k++;
                    }
                }
            }

            return result;
        }

        public string[] SubStringArray(string[] myStringArray, int[] myNumberList)
        {
            string[] result = new string[myStringArray.Length - myNumberList.Length];
            int startNumber = 0;  //for 中奖名单不含抽奖名单的第一个元素
            int endNumber = 0;  //for 中奖名单含抽奖名单的第一个元素

            if ((myNumberList[myNumberList.Length - 1] - myNumberList[0] == myNumberList.Length - 1) && myNumberList[0] != 0)  //中奖名单不含抽奖名单的第一个元素
            {
                startNumber = myNumberList[0];

                for (int i = 0; i < result.Length; i++)
                {
                    if (i < startNumber)
                    {
                        result[i] = myStringArray[i];
                    }
                    else
                    {
                        result[i] = myStringArray[i + myNumberList.Length];
                    }
                }
            }
            else //中奖名单含抽奖名单的第一个元素
            {
                endNumber = myNumberList[myNumberList.Length - 1];

                for (int i = 0; i < result.Length; i++)
                {
                    result[i] = myStringArray[endNumber + 1 + i];
                }
            }

            return result;
        }
        #endregion

        #region 小游戏 下一奖项
        private void NextLittleGame_Click(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                LittleGamePrizeName.Text = "最佳烧脑奖";
            }));

            isSix = true;
            gameCount = 0;

            NextLittleGame.IsEnabled = false;
            StartLittleGame.IsEnabled = true;

            LittleGameNameResult.Text = "";
            LittleGameNameFlow.Text = "抽奖箱";

        }
        #endregion

        #endregion

        #region 页面处理
        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            littleGameTimer.Elapsed += new ElapsedEventHandler(littleGameTimerHandle);
            NextLittleGame.IsEnabled = false;
            luckyDrawTimer.Elapsed += new ElapsedEventHandler(luckyDrawTimerHandle);
            NextLuckyDraw.IsEnabled = false;
            timeDelay.Elapsed += new ElapsedEventHandler(TimeDelayHandle);
            luckyDrawOnePrizeTimer.Elapsed += new ElapsedEventHandler(luckyDrawOnePrizeTimerHandle);

        }


        private void myWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("你要退出软件吗？", "问询", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if(result!=MessageBoxResult.Yes)
            {
                e.Cancel = true;
            }
        }

        #endregion

        #region 抽奖
        System.Timers.Timer luckyDrawTimer = new System.Timers.Timer(100);
        System.Timers.Timer luckyDrawOnePrizeTimer = new System.Timers.Timer(100);
        int luckyDrawPrize = 5;  //奖项名称
        int luckyDrawPrizeTime = 0;  //抽奖次数
        int luckyDrawNameListNumber = 0;  //中奖序号

        string[] SubRandomThreeFourFivePrize = new string[1];
        string[] SubRandomOneTwoPrize = new string[1];

//        string[][] PresentName = new string[5][]
//        {
//            new string[]{"ipone XS 手机"},
//        new string[]{"Apple Macbook Air 13.3英寸笔记本 i5/8G/128G","Song 数码相机（RX100M5A) +拍摄手柄+相机包"},
//        new string[]{"大疆灵眸口袋云台相机套装","Apple watch series 4(黑色 GPS款 40mm)","Apple iPad 平板电脑 9.7英寸（128G WLAN版）银色"},
//        new string[]{"Kindle(第四代 6英寸 8GB）","双立人如意刀具9件套","米家(MIJIA)压力IH电饭煲 小米电饭锅 3L","海尔扫地机器人"},
//        new string[]{"华为备咖存储", "华为备咖存储" ,"JBL 无线蓝牙耳机", "JBL 无线蓝牙耳机" ,"华为畅享8e"}
//};

        string[] PresentLevel = new string[5] { "一等奖", "二等奖", "三等奖", "四等奖", "五等奖"};

        System.Timers.Timer timeDelay = new System.Timers.Timer(10);

        int[] onePrizeCountArray = new int[]
        {
            100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,
            200,300,400,500,600,700,800,900,1000,2000,
            1000,900,800,700,600,500,400,300,200,
            100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,
            1000,5000,
            1000,500,200,
            100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,
            500,1000,3000,
            1000,200,
            100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,
            200,300,400,500,600,700,800,900,1000,
            10000,10000,10000,10000,10000,

        };

        int onePrizeCount = 0;

        #region 抽奖 开始/停止
        private void StartLuckyDraw_Click(object sender, RoutedEventArgs e)
        {            
            if (StartLuckyDraw.Content.ToString()=="开始")
            {
                StartLuckyDraw.Content = "停止";

                if(luckyDrawPrize==2&&luckyDrawPrizeTime==0)
                {
                    luckyDrawNameListNumber = 0;
                }
                else
                {
                    if(luckyDrawNameListNumber!=0)
                    {
                        luckyDrawNameListNumber--;
                    }
                }

                if(luckyDrawPrize==1&& luckyDrawPrizeTime==0)
                {
                    StartLuckyDraw.IsEnabled = false;
                    SpeedAdjust.Visibility = Visibility.Collapsed;
                    Slow.Visibility = Visibility.Collapsed;
                    Fast.Visibility = Visibility.Collapsed;
                    SpeedShow.Visibility = Visibility.Collapsed;
                   

                    onePrizeCount = 0;

                    luckyDrawOnePrizeTimer.Start();

                }
                else
                {
                    luckyDrawTimer.Start();
                }

            }
            else if(StartLuckyDraw.Content.ToString() == "停止")
            {
                StartLuckyDraw.Content = "开始";
                luckyDrawTimer.Stop();

                StartLuckyDraw.IsEnabled = false;
                timeDelay.Start();

                //五等奖处理
                if (luckyDrawPrize==5)
                {
                    if(luckyDrawPrizeTime==0)
                    {
                        SubRandomThreeFourFivePrize = new string[RandomThreeFourFivePrize.Length - 1];
                        SubRandomThreeFourFivePrize = LuckyDrawGeneratorNewNameList(RandomThreeFourFivePrize, luckyDrawNameListNumber);

                        luckyDrawPrizeTime++;
                    }
                    else if(luckyDrawPrizeTime>0&&luckyDrawPrizeTime<4)
                    {
                        GetSubRandomThreeFourFivePrize();

                        luckyDrawPrizeTime++;
                    }
                    else if(luckyDrawPrizeTime==4)
                    {
                        GetSubRandomThreeFourFivePrize();

                        luckyDrawPrizeTime++;

                        StartLuckyDraw.IsEnabled = false;
                        NextLuckyDraw.IsEnabled = true;
                    }

                    //ExportExcelSub(SubRandomThreeFourFivePrize);

                }

                //四等奖处理
                if (luckyDrawPrize == 4)
                {
                    if (luckyDrawPrizeTime >= 0 && luckyDrawPrizeTime < 3)
                    {
                        GetSubRandomThreeFourFivePrize();

                        luckyDrawPrizeTime++;
                    }
                    else if (luckyDrawPrizeTime == 3)
                    {
                        GetSubRandomThreeFourFivePrize();
                        luckyDrawPrizeTime++;

                        StartLuckyDraw.IsEnabled = false;
                        NextLuckyDraw.IsEnabled = true;
                    }

                    //ExportExcelSub(SubRandomThreeFourFivePrize);

                }

                //三等奖处理
                if (luckyDrawPrize == 3)
                {
                    if (luckyDrawPrizeTime >= 0 && luckyDrawPrizeTime < 2)
                    {
                        GetSubRandomThreeFourFivePrize();

                        luckyDrawPrizeTime++;

                        //ExportExcelSub(SubRandomThreeFourFivePrize);
                    }
                    else if (luckyDrawPrizeTime == 2)
                    {
                        luckyDrawPrizeTime++;


                        StartLuckyDraw.IsEnabled = false;
                        NextLuckyDraw.IsEnabled = true;
                    }
                }

                //二等奖处理
                if (luckyDrawPrize == 2)
                {
                    if (luckyDrawPrizeTime == 0)
                    {
                        SubRandomOneTwoPrize = new string[GetNewNameListLength(RandomOneTwoPrize, luckyDrawNameListNumber)];
                        SubRandomOneTwoPrize = LuckyDrawGeneratorNewNameListForOneTwoPrize(RandomOneTwoPrize, luckyDrawNameListNumber);

                        luckyDrawPrizeTime++;

                    }                    
                    else if (luckyDrawPrizeTime == 1)
                    {
                        string[] tempStringArray = new string[GetNewNameListLength(SubRandomOneTwoPrize, luckyDrawNameListNumber)];
                        tempStringArray = LuckyDrawGeneratorNewNameListForOneTwoPrize(SubRandomOneTwoPrize, luckyDrawNameListNumber);

                        SubRandomOneTwoPrize = new string[tempStringArray.Length];
                        for (int i = 0; i < tempStringArray.Length; i++)
                        {
                            SubRandomOneTwoPrize[i] = tempStringArray[i];
                        }


                        luckyDrawPrizeTime++;


                        StartLuckyDraw.IsEnabled = false;
                        NextLuckyDraw.IsEnabled = true;
                    }

                    //ExportExcelSub(SubRandomOneTwoPrize);

                }

                ////一等奖处理
                //if (luckyDrawPrize == 1)
                //{
                //    if (luckyDrawPrizeTime == 0)
                //    {
                //        luckyDrawPrizeTime++;

                //        StartLuckyDraw.IsEnabled = false;
                //        NextLuckyDraw.IsEnabled = false;

                //        fireworks.Visibility = Visibility.Visible;
                //    }
                //}
            }
        }

        //获得三、四、五等奖的子名单
        public void GetSubRandomThreeFourFivePrize()
        {
            string[] tempStringArray = new string[SubRandomThreeFourFivePrize.Length - 1];
            tempStringArray = LuckyDrawGeneratorNewNameList(SubRandomThreeFourFivePrize, luckyDrawNameListNumber);

            SubRandomThreeFourFivePrize = new string[tempStringArray.Length];
            for (int i = 0; i < tempStringArray.Length; i++)
            {
                SubRandomThreeFourFivePrize[i] = tempStringArray[i];
            }
        }

        private void SpeedAdjust_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            luckyDrawTimer.Interval = SpeedAdjust.Value;
        }


        private void luckyDrawTimerHandle(object source, ElapsedEventArgs e)
        {
            if (luckyDrawPrize == 5)
            {
                if (luckyDrawPrizeTime == 0)
                {
                    LuckyDrawFlow(RandomThreeFourFivePrize);
                }
                else 
                {
                    LuckyDrawFlow(SubRandomThreeFourFivePrize);
                }
            }
            else if(luckyDrawPrize==4|| luckyDrawPrize == 3)
            {
                LuckyDrawFlow(SubRandomThreeFourFivePrize);
            }            
            else if (luckyDrawPrize == 2)
            {
                if (luckyDrawPrizeTime == 0)
                {
                    LuckyDrawFlow(RandomOneTwoPrize);
                }
                else
                {
                    LuckyDrawFlow(SubRandomOneTwoPrize);
                }
            }
            //else if(luckyDrawPrize==1)
            //{
            //    if (luckyDrawPrizeTime == 0)
            //    {
            //        LuckyDrawFlow(SubRandomOneTwoPrize);
            //    }
               
            //}
        }

        private void luckyDrawOnePrizeTimerHandle(object source, ElapsedEventArgs e)
        {
            LuckyDrawFlow(SubRandomOneTwoPrize);

            if (onePrizeCount<onePrizeCountArray.Length)
            {
                luckyDrawOnePrizeTimer.Interval = onePrizeCountArray[onePrizeCount++];
            }
            else
            {
                luckyDrawOnePrizeTimer.Stop();

                luckyDrawPrizeTime++;

                Thread.Sleep(8000);

                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    StartLuckyDraw.IsEnabled = false;
                    NextLuckyDraw.IsEnabled = false;

                    fireworks.Visibility = Visibility.Visible;
                }));               
            }
        }


            public void LuckyDrawFlow(string[] myStringArray)
        {
            if (luckyDrawNameListNumber >= myStringArray.Length)
            {
                luckyDrawNameListNumber = 0;
            }
           
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                LuckyDrawNameFlow.Text = myStringArray[luckyDrawNameListNumber++];

            }));           
            
        }


        public string[] LuckyDrawGeneratorNewNameList(string[] myStringArray, int elementNumber)
        {
            string[] result = new string[myStringArray.Length - 1];

            if(elementNumber==0)
            {
                for(int i=0;i<result.Length;i++)
                {
                    result[i] = myStringArray[i + 1];
                }
            }
            else
            {
                for(int i=0;i<result.Length;i++)
                {
                    if(i<elementNumber-1)
                    {
                        result[i] = myStringArray[i];
                    }
                    else
                    {
                        result[i] = myStringArray[i + 1];
                    }
                }
            }

            return result;
        }

        public string[] LuckyDrawGeneratorNewNameListForOneTwoPrize(string[] myStringArray, int elementNumber)
        {
            string LuckyName = "";

            if (elementNumber==0)
            {
                LuckyName = myStringArray[elementNumber];
            }
            else
            {
                LuckyName = myStringArray[elementNumber - 1];
            }

            ArrayList myArrayList = new ArrayList();

            for(int i=0;i<myStringArray.Length;i++)
            {
                if(myStringArray[i]==LuckyName)
                {
                    myArrayList.Add(i);
                }
            }

            string[] result = new string[myStringArray.Length - myArrayList.Count];

            int k = 0;
            int z = 0;
            for(int i=0;i<myStringArray.Length;i++)
            {
                if(i==Convert.ToInt32( myArrayList[k]))
                {
                    if(k<myArrayList.Count-1)
                    {
                        k++;
                    }
                }
                else
                {
                    result[z++] = myStringArray[i];
                }
            }

            return result;
        }

        public int GetNewNameListLength(string[] myStringArray, int elementNumber)
        {
            string LuckyName = myStringArray[elementNumber];
            ArrayList myArrayList = new ArrayList();

            for (int i = 0; i < myStringArray.Length; i++)
            {
                if (myStringArray[i] == LuckyName)
                {
                    myArrayList.Add(i);
                }
            }

            return myStringArray.Length - myArrayList.Count;
        }

        private void TimeDelayHandle(object source, ElapsedEventArgs e)
        {
            if((luckyDrawPrize==5&&luckyDrawPrizeTime==5)|| (luckyDrawPrize == 4 && luckyDrawPrizeTime == 4)||(luckyDrawPrize == 3 && luckyDrawPrizeTime == 3)|| (luckyDrawPrize == 2 && luckyDrawPrizeTime == 2)|| (luckyDrawPrize == 1 && luckyDrawPrizeTime == 1))
            {
                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    StartLuckyDraw.IsEnabled = false;
                }));
            }
            else
            {
                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    StartLuckyDraw.IsEnabled = true;
                }));
            }

            timeDelay.Stop();
        }
        #endregion

        #region 抽奖 下一奖项

        private void NextLuckyDraw_Click(object sender, RoutedEventArgs e)
        {
            luckyDrawPrize--;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                LuckyDrawPrizeName.Text = PresentLevel[luckyDrawPrize - 1];
                LuckyDrawNameFlow.Text = "抽奖箱";

            }));

            luckyDrawPrizeTime = 0;

            NextLuckyDraw.IsEnabled = false;
            StartLuckyDraw.IsEnabled = true;

        }


        #endregion

        #endregion      

        #region 设置
        private void SettingConfirm_Click(object sender, RoutedEventArgs e)
        {
            if(SettingOnePrize.IsChecked==true)
            {
                NextLuckyDraw.IsEnabled = false;
                StartLuckyDraw.IsEnabled = true;

                luckyDrawPrize = 1;
                luckyDrawPrizeTime = 0;


                luckyDrawOnePrizeTimer.Interval = 100;

                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    LuckyDrawPrizeName.Text = PresentLevel[luckyDrawPrize - 1];
                    LuckyDrawNameFlow.Text = "抽奖箱";
                    StartLuckyDraw.Content = "开始";

                }));

                SubRandomOneTwoPrize = new string[RandomOneTwoPrize.Length];
                for(int i=0;i<SubRandomOneTwoPrize.Length;i++)
                {
                    SubRandomOneTwoPrize[i] = RandomOneTwoPrize[i];
                }


                luckyDrawNameListNumber = new Random().Next(0, SubRandomOneTwoPrize.Length);



                fireworks.Visibility = Visibility.Collapsed;

                SpeedAdjust.Visibility = Visibility.Collapsed;
                Slow.Visibility = Visibility.Collapsed;
                Fast.Visibility = Visibility.Collapsed;
                SpeedShow.Visibility = Visibility.Collapsed;
                

                MessageBox.Show("设置成功！");

                LuckyDraw.IsSelected = true;

            }

            if (SettingTwoPrize.IsChecked == true)
            {
                NextLuckyDraw.IsEnabled = false;
                StartLuckyDraw.IsEnabled = true;

                luckyDrawPrize = 2;
                luckyDrawPrizeTime = 0;

                luckyDrawNameListNumber = 0;

                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    LuckyDrawPrizeName.Text = PresentLevel[luckyDrawPrize - 1];
                    LuckyDrawNameFlow.Text = "抽奖箱";

                    StartLuckyDraw.Content = "开始";


                }));

                fireworks.Visibility = Visibility.Collapsed;

                SpeedAdjust.Visibility = Visibility.Visible;
                Slow.Visibility = Visibility.Visible;
                Fast.Visibility = Visibility.Visible;
                SpeedShow.Visibility = Visibility.Visible;
                

                MessageBox.Show("设置成功！");

                LuckyDraw.IsSelected = true;

            }

            if (SettingThreePrize.IsChecked == true)
            {
                NextLuckyDraw.IsEnabled = false;
                StartLuckyDraw.IsEnabled = true;

                luckyDrawPrize = 3;
                luckyDrawPrizeTime = 0;

                luckyDrawNameListNumber = 0;

                SubRandomThreeFourFivePrize = new string[RandomThreeFourFivePrize.Length];
                for (int i = 0; i < SubRandomThreeFourFivePrize.Length; i++)
                {
                    SubRandomThreeFourFivePrize[i] = RandomThreeFourFivePrize[i];
                }

                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    LuckyDrawPrizeName.Text = PresentLevel[luckyDrawPrize - 1];
                    LuckyDrawNameFlow.Text = "抽奖箱";

                    StartLuckyDraw.Content = "开始";


                }));

                fireworks.Visibility = Visibility.Collapsed;


                SpeedAdjust.Visibility = Visibility.Visible;
                Slow.Visibility = Visibility.Visible;
                Fast.Visibility = Visibility.Visible;
                SpeedShow.Visibility = Visibility.Visible;
                

                MessageBox.Show("设置成功！");

                LuckyDraw.IsSelected = true;

            }

            if (SettingFourPrize.IsChecked == true)
            {
                NextLuckyDraw.IsEnabled = false;
                StartLuckyDraw.IsEnabled = true;

                luckyDrawPrize = 4;
                luckyDrawPrizeTime = 0;

                luckyDrawNameListNumber = 0;

                SubRandomThreeFourFivePrize = new string[RandomThreeFourFivePrize.Length];
                for (int i = 0; i < SubRandomThreeFourFivePrize.Length; i++)
                {
                    SubRandomThreeFourFivePrize[i] = RandomThreeFourFivePrize[i];
                }

                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    LuckyDrawPrizeName.Text = PresentLevel[luckyDrawPrize - 1];
                    LuckyDrawNameFlow.Text = "抽奖箱";

                    StartLuckyDraw.Content = "开始";


                }));

                fireworks.Visibility = Visibility.Collapsed;

                SpeedAdjust.Visibility = Visibility.Visible;
                Slow.Visibility = Visibility.Visible;
                Fast.Visibility = Visibility.Visible;
                SpeedShow.Visibility = Visibility.Visible;
               

                MessageBox.Show("设置成功！");

                LuckyDraw.IsSelected = true;

            }

            if (SettingFivePrize.IsChecked == true)
            {
                NextLuckyDraw.IsEnabled = false;
                StartLuckyDraw.IsEnabled = true;

                luckyDrawPrize = 5;
                luckyDrawPrizeTime = 0;

                luckyDrawNameListNumber = 0;

                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    LuckyDrawPrizeName.Text = PresentLevel[luckyDrawPrize - 1];
                    LuckyDrawNameFlow.Text = "抽奖箱";

                    StartLuckyDraw.Content = "开始";


                }));

                fireworks.Visibility = Visibility.Collapsed;

                SpeedAdjust.Visibility = Visibility.Visible;
                Slow.Visibility = Visibility.Visible;
                Fast.Visibility = Visibility.Visible;
                SpeedShow.Visibility = Visibility.Visible;
                

                MessageBox.Show("设置成功！");

                LuckyDraw.IsSelected = true;

            }

            if (SettingSixPrize.IsChecked == true)
            {
                NextLittleGame.IsEnabled = false;
                StartLittleGame.IsEnabled = true;

                isSix = true;
                gameCount = 0;
                nameListNumber = 0;                

                subRandomSixSevenPrize = new string[RandomSixSevenPrize.Length];
                for (int i = 0; i < subRandomSixSevenPrize.Length; i++)
                {
                    subRandomSixSevenPrize[i] = RandomSixSevenPrize[i];
                }

                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    LittleGamePrizeName.Text = "最佳烧脑奖";
                    LittleGameNameFlow.Text = "抽奖箱";
                    LittleGameNameResult.Text = "抽中名单";

                }));

                MessageBox.Show("设置成功！");

                LittleGame.IsSelected = true;
            }

            if (SettingSevenPrize.IsChecked == true)
            {
                NextLittleGame.IsEnabled = false;
                StartLittleGame.IsEnabled = true;

                isSix = false;
                gameCount = 0;
                nameListNumber = 0;

                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    LittleGamePrizeName.Text = "最有希望奖";
                    LittleGameNameFlow.Text = "抽奖箱";
                    LittleGameNameResult.Text = "抽中名单";


                }));

                MessageBox.Show("设置成功！");

                LittleGame.IsSelected = true;

            }

        }

        #endregion       
    }

    public class ExcelHelper : IDisposable
    {
        private string fileName = null; //文件名
        private IWorkbook workbook = null;
        private FileStream fs = null;
        private bool disposed;
        public ExcelHelper(string fileName)//构造函数，读入文件名
        {
            this.fileName = fileName;
            disposed = false;
        }
        /// 将excel中的数据导入到DataTable中
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public System.Data.DataTable ExcelToDataTable(string sheetName, bool isFirstRowColumn)
        {
            ISheet sheet = null;
            System.Data.DataTable data = new System.Data.DataTable();
            int startRow = 0;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                workbook = WorkbookFactory.Create(fs);
                if (sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    if (sheet == null)
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                }
                else
                {
                    sheet = workbook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号，即总的列数
                    if (isFirstRowColumn)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null)
                                {
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;//得到项标题后
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }
                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　　　　　　

                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                                dataRow[j] = row.GetCell(j).ToString();
                        }
                        data.Rows.Add(dataRow);
                    }
                }
                return data;
            }
            catch (Exception ex)//打印错误信息
            {
                MessageBox.Show("Exception: " + ex.Message);
                return null;
            }
        }

        //将DataTable数据导入到excel中
        //<param name="data">要导入的数据</param>
        //<param name="sheetName">要导入的excel的sheet的名称</param>
        //<param name="isColumnWritten">DataTable的列名是否要导入</param>
        //<returns>导入数据行数(包含列名那一行)</returns>
        public int DataTableToExcel(System.Data.DataTable data, string sheetName, bool isColumnWritten)
        {
            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;

            fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                workbook = new XSSFWorkbook();
            else if (fileName.IndexOf(".xls") > 0) // 2003版本
                workbook = new HSSFWorkbook();

            try
            {
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet(sheetName);
                }
                else
                {
                    return -1;
                }

                if (isColumnWritten == true) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                    }
                    count = 1;
                }
                else
                {
                    count = 0;
                }

                for (i = 0; i < data.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }
                workbook.Write(fs); //写入到excel
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return -1;
            }
        }

        public void Dispose()//IDisposable为垃圾回收相关的东西，用来显式释放非托管资源,这部分目前还不是非常了解
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    if (fs != null)
                        fs.Close();
                }
                fs = null;
                disposed = true;
            }
        }
    }

}
