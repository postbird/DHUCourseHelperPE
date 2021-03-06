﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Reflection;

namespace DHUTeacerCourseHelper
{
    public partial class Form1 : Form
    {
        private string courseFilePath = "";
        private string courseCodeFilePath = "";
        private string courseExtention;
        private string courseCodeTable;
        private string courseCodeExtention;
        private DataSet courseCodeDs=new DataSet();
        private ArrayList hunClass = new ArrayList();
        private ArrayList extManClass = new ArrayList();//主要记录没有 男  这个标记的一些课程

        public Form1()
        {
            InitializeComponent();
            statusLabel.Text = "准备就绪.Powered by Postbird.";
            initRichText();

        }

        private void courseFileButton_Click(object sender, EventArgs e)
        {
            //初始化一个openfileDialog
            OpenFileDialog fileDialog = new OpenFileDialog();
            //设置过滤属性 xls xlsx
            fileDialog.Filter = "Office 2007后版本(*.xlsx)|*.xlsx|Office 2003版本(*.xls)|*.xls";
            //判断用户是否选择了文件
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                //获取用户选择文件的后缀名
                string extension = Path.GetExtension(fileDialog.FileName);
                this.courseExtention = extension;
                //声明允许的后缀名
                string[] str = new string[] { ".xls", ".xlsx" };
                if (!((IList)str).Contains(extension))
                {
                    MessageBox.Show("只能上传后缀为 .xls | .xlsx 的文件");
                    courseFileTextBox.Text = "";
                }
                else
                {
                    courseFileTextBox.Text = fileDialog.FileName;
                    this.courseFilePath = fileDialog.FileName;
                }
            }
        }

        private void courseCodeFileButton_Click(object sender, EventArgs e)
        {
            //初始化一个openfileDialog
            OpenFileDialog fileDialog = new OpenFileDialog();
            //设置过滤属性 xls xlsx
            fileDialog.Filter = "Office 2007后版本(*.xlsx)|*.xlsx|Office 2003版本(*.xls)|*.xls";
            //判断用户是否选择了文件
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                //获取用户选择文件的后缀名
                string extension = Path.GetExtension(fileDialog.FileName);
                this.courseCodeExtention = extension;
                //声明允许的后缀名
                string[] str = new string[] { ".xls", ".xlsx" };
                if (!((IList)str).Contains(extension))
                {
                    MessageBox.Show("只能上传后缀为 .xls | .xlsx 的文件");
                    courseCodeFileTextBox.Text = "";
                }
                else
                {
                    courseCodeFileTextBox.Text = fileDialog.FileName;
                    this.courseCodeFilePath = fileDialog.FileName;
                }

            }
        }
        //初始化 initRichText
        private void initRichText()
        {
            string tmpRichText = "";
            int flag = 0;
            //男生班
            extManClass = new ArrayList();
            extManClass.Add("散打");
            for (int i = 0; i < extManClass.Count; i++)
            {
                if (flag > 0)
                {
                    tmpRichText += System.Environment.NewLine;
                }
                tmpRichText += extManClass[i];
                flag++;
            }
            //输出
            manClassRichTextBox.Text = tmpRichText;

            //混合班
            hunClass = new ArrayList();
            hunClass.Add("保健班");
            hunClass.Add("旱冰");
            hunClass.Add("棒球");
            hunClass.Add("野外");
            hunClass.Add("拓展");
            hunClass.Add("交谊舞");
            hunClass.Add("交谊舞2");
            hunClass.Add("高尔夫");
            hunClass.Add("定向");
            hunClass.Add("游泳");
            hunClass.Add("游泳(高)");
            hunClass.Add("救生");
            hunClass.Add("冰壶"); 
            hunClass.Add("滑冰");
            hunClass.Add("跆拳道");
            for (int i = 0; i < hunClass.Count; i++)
            {
                if (flag>0)
                {
                    tmpRichText += System.Environment.NewLine;
                }
                tmpRichText += hunClass[i];
                flag++;
            }
            //输出
            hunClassRichText.Text = tmpRichText;
        }
        //重置
        private void resetButton_Click(object sender, EventArgs e)
        {
            courseFilePath = "";
            courseCodeFilePath = "";
            courseExtention="";
            courseCodeTable="";
            courseCodeExtention="";
            courseCodeDs = new DataSet();
            courseCodeFileTextBox.Text = "";
            courseFileTextBox.Text = "";
            statusLabel.Text = "准备就绪.Powered by Postbird.";
            statusCountTextBox.Text = "";
            initRichText();

        }
        //处理
        private void submitButton_Click(object sender, EventArgs e)
        {
            //设置按钮的不可用性
            courseFileButton.Enabled = false;
            courseCodeFileButton.Enabled = false;
            resetButton.Enabled = false;
            submitButton.Enabled = false;

            //判断是否为空 
            if (courseFileTextBox.Text.ToString().Trim().Length == 0 || courseCodeFileTextBox.Text.ToString().Trim().Length == 0)
            {
                MessageBox.Show("文件不能为空!");
                //恢复按钮的可用性
                courseFileButton.Enabled = true;
                courseCodeFileButton.Enabled = true;
                resetButton.Enabled = true;
                submitButton.Enabled = true;
            }
            else
            {
                /*****************************************************************

                           获取课程代码部分 使用oledb方式进行操作
               *****************************************************************/

                //获取文本内容
                courseCodeFilePath = courseCodeFileTextBox.Text.ToString().Trim();
                //创建连接 以便发生异常时关闭连接(建在try外)
                OleDbConnection courseCodeOleConn = new OleDbConnection();
                try
                {
                    //判断文件 2003还是2007 分别创建文件链接
                    //创建连接，引用协议
                    string courseCodeConn = "";
                    if (this.courseCodeExtention == ".xls")
                    {
                        //2003（Microsoft.Jet.Oledb.4.0）
                        courseCodeConn = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", this.courseCodeFilePath);
                    }
                    else
                    {
                        //2010（Microsoft.ACE.OLEDB.12.0）
                        courseCodeConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'", this.courseCodeFilePath);
                    }

                    //打开连接并执行sql语句，末尾需要关闭连接
                    courseCodeOleConn = new OleDbConnection(courseCodeConn);
                    courseCodeOleConn.Open();
                    //获取所有的表 默认处理第一张表
                    System.Data.DataTable courseTables = courseCodeOleConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    //获取第一张表的名字用于查询
                    this.courseCodeTable = courseTables.Rows[0]["TABLE_NAME"].ToString();
                    // 输出状态
                    statusLabel.Text = "从" + this.courseCodeTable + "表读取课程代码...";
                    //执行sql查询功能,保存到dataset中    
                    String sql = string.Format("SELECT * FROM  [{0}]", this.courseCodeTable);
                    //创建查询语句
                    OleDbDataAdapter oleAdapter = new OleDbDataAdapter(sql, courseCodeOleConn);
                    //创建dataSet保存数据
                    DataSet ds = new DataSet();
                    //获得数据
                    oleAdapter.Fill(ds, this.courseCodeTable);
                    //对创建dataSet保存数据的遍历
                    //将中文括号变成英文括号 方便匹配
                    /*
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        for (int j = 2; j < ds.Tables[0].Columns.Count; j++)
                        {
                            //MessageBox.Show(ds.Tables[0].Rows[i][j].ToString(), "提示框");
                            ds.Tables[0].Rows[i][j] = ds.Tables[0].Rows[i][j].ToString();
                            ds.Tables[0].Rows[i][j]=ds.Tables[0].Rows[i][j].ToString();
                        }
                    }
                    */
                    //返回处理后的ds
                    this.courseCodeDs = ds;

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "发生错误");
                }
                finally
                {
                    //关闭连接
                    courseCodeOleConn.Close();
                }

                /*****************************************************************

                                  操作排课表  使用Cells[x, y]

                *****************************************************************/
                // 输出状态
                statusLabel.Text = "处理排课...";
                try
                {
                    //声明app
                    Microsoft.Office.Interop.Excel.Application courseApp = new Microsoft.Office.Interop.Excel.Application();
                    //让后台执行设置为不可见
                    courseApp.Visible = false;
                    Workbooks wbks = courseApp.Workbooks;
                    //获取文档
                    _Workbook _wbk = wbks.Add(courseFilePath);
                    //获取表
                    Sheets shs = _wbk.Sheets;
                    //处理每个表  因为只有松江区校区和延安路校区 因此这里就写死了
                    // 只有两个sheet 需要处理分别是1 2  没有0 非常蛋疼。。。
                    //同样 在判断大一和大二的问题上 由于排课表是确定形式的 因此B列存在的是大一还是大二的问题
                    //[废弃]  而A列是上课的时间 因此也写死了 不去遍历 最高到L 同时 最长到34  
                    //[废弃]  因此在处理的过程中 直接处理的是 C-L 以及 1-40 这样子写死遍历
                    //[更新]  本来打算使用B1来获取数据，但是发现存在合并的单元格，因此最后还是采用[1,2]的方式获取数据
                    //[更新]  因此在本来 C-L 的循环变成了 3-12的循环 还是没有0
                    //[更新]  同时循环的方式也进行了改变 不过 在判断是否是大一的问题上  
                    //可以判断值是否为空 进行跳过

                    //记录最后的统计  松江校区:共x 男 12 女 12 混 12  延安路校区:共x 男 12 女 123 混12
                    string statusCountText = "";

                    for (int m = 1; m <= 2; m++)
                    {
                        /*******************************************************************************************
                         * 
                         *    //需要记录男生多少课  女生多少课 混合班有多少
                         *    2016-11-17 更新 同时把大一和大二/3/4的分别统计出来
                         *      
                         * *************************************************************************************/
                        int oneMenCount = 0;//大一 男
                        int oneWomenCount = 0;//大一 女
                        int oneHunCount = 0;//大一 混

                        int menCount = 0;//大二/3/4 男
                        int womenCount = 0;//大二/3/4 女
                        int hunCount = 0;//大二/3/4 混
                        //简单的循环处理 1代表松江 2代表延安路校区
                        string statusText = "";
                        
                        _Worksheet tmpTable = (_Worksheet)shs.get_Item(m);
                        if (m == 1)
                        {
                            statusText = "松江校区";//用于输出状态
                        }
                        else
                        {
                            statusText = "延安路校区";

                        }
                        // 输出状态
                        statusLabel.Text = "处理" + statusText + "排课...";
                        //记录处理次数
                        int courseCount = 0;
                        //循环处理 松江校区排课的sheet
                        for (int j = 3; j <= 12; j++)//列号
                        {
                            //行号
                            for (int i = 1; i <= 40; i++)
                            {
                                //改进
                                Range curentCell = (Range)tmpTable.Cells[i, j];
                                string tmpText = curentCell.Text.ToString().Trim();  //单元格文本
                                                                                     // MessageBox.Show(i+" "+j+" "+tmpText);
                                if (tmpText.Equals(""))
                                {
                                    continue;
                                }

                                //将中文括号变成英文括号

                                //针对课程代码中的数据进行匹配 用了全匹配 可能比较慢---贼慢

                                //解释在这里如果不懂啥意思 可以看看下面再上来
                                // 发现 存的时候 是 王:男篮(高) 因此是不能直接相等 因此需要去掉前两个 再来一个临时字符
                                string tmpCouse = "";
                                if (tmpText.Length >= 2)
                                {
                                    tmpText = tmpText.Replace(" ", "");
                                    tmpCouse = tmpText.Substring(2);
                                }
                                //  MessageBox.Show(tmpCouse);
                                /******************************************************************************************

                                                    加的处理男生和女生以及混合课程的统计量

                                        如果可以匹配到课程代码 就进行男女的匹配 应为男的一般都有男这个字 因此只要包含男这个字就算是男生班
                                        而混合班比较少 因此使用richTextBox进行内容的获取，由此得知混合班有哪些
                                         如果两个都无法匹配 就算成女生版 因为女生班名字并没有比较明显的标志
                                 *********************************************************************************************/
                                //获取混合课程内容
                                string[] hunClassStrArray = hunClassRichText.Text.ToString().Split('\r', '\n'); 
                                //2016-11-17 获取没有带男子的课程内容
                                string[] manClassStrArray = manClassRichTextBox.Text.ToString().Split('\r', '\n');

                                //去掉所有的空格
                                for (int t = 0; t < hunClassStrArray.Length; t++)
                                {
                                    //MessageBox.Show(hunClassStrArray[t]);
                                    hunClassStrArray[t] = hunClassStrArray[t].Trim();
                                }

                                //*********************************匹配
                                for (int k = 0; k < this.courseCodeDs.Tables[0].Rows.Count; k++)
                                {
                                    //课程代码结构很明显  从 0 开始 0-序号 1 代表课程名称 2代表新生 3代表老生代码
                                    //[废弃]如果课程中包含了课程代码中的文字 那么进行匹配
                                    //[废弃]针对有些时候没有写男女  加了一次判定 本列第三行 是代表了男生女生
                                    //[废弃][第一次更新]第二次判断 如果里面没有写明男女 则判断课程代码中包含 比如 攀岩（男）【代码】 而课程中是 攀岩 列标题写了男 这样子也可以匹配
                                    //[废弃][第一次更新]上面的做法容易产生 男 女 单个字符的匹配 因此加了个判定不能只是 男 女
                                    //[第二次更新] 发现如果排课中 字数多 包含 字数少 不能正确匹配 比如 男篮(高)最后匹配的是 男篮。
                                    //[第二次更新] 只要一个条件 那就是相等！！就是强制让你一样，不服咬我啊！ 判断是否完全相等
                                    // 日 le dog  看for循环上面的解释 ||

                                    // 获取临时代码
                                    string tmpCourseCode = this.courseCodeDs.Tables[0].Rows[k][1].ToString();
                                    if (tmpCouse.Equals(tmpCourseCode))
                                    {
                                        //判断是新生还是老生
                                        //根据上面写死的 其中 当前行的第2列是表示大一还是大二、三的
                                        //因此根据 大一来判断新生否则是老生
                                        Range tmpGradeRange = (Range)tmpTable.Cells[i, 2];
                                        string tmpGrade = "";
                                        //判断是否是合并单元格
                                        if ((bool)tmpGradeRange.MergeCells)
                                        {
                                            Range mergeArea = (Range)tmpTable.Cells[tmpGradeRange.MergeArea.Row, tmpGradeRange.MergeArea.Column];
                                            tmpGrade = mergeArea.Text.ToString();
                                        }
                                        else
                                        {
                                            tmpGrade = tmpGradeRange.Text.ToString();
                                        }

                                        /**********************************************************************************
                                         * 
                                         *      这里是处理男生 女生 混合班的部分
                                         *      
                                         *      2016-11-17  我把它放进了处理大一和大二三里面去了
                                         * 
                                         ***********************************************************/
                                        
                                        //把课程代码加上去  大一在课程代码中是新生 
                                        if (tmpGrade.Equals("大一"))
                                        {
                                            tmpText = tmpText + "      " + this.courseCodeDs.Tables[0].Rows[k][2];
                                            
                                            //进行大一的统计
                                            int courseFlag = 0;//匹配标志  
                                            if (tmpCouse.Contains("男"))
                                            {
                                                oneMenCount++;
                                                courseFlag = 1;

                                            }
                                            //[废弃]不是男生
                                            //2016-11-17 还需要匹配一次textbox里面的男生班 没有带男字标志的
                                            else
                                            {
                                                //没有男字标志的男生班
                                                
                                                for (int t = 0; t < manClassStrArray.Length; t++)
                                                {
                                                    if (manClassStrArray[t] == tmpCouse)
                                                    {
                                                        oneMenCount++;
                                                        courseFlag = 1;
                                                        break;
                                                    }
                                                }
                                                //混合
                                                if (courseFlag == 0)
                                                {
                                                    for (int t = 0; t < hunClassStrArray.Length; t++)
                                                    {
                                                        if (hunClassStrArray[t] == tmpCouse)
                                                        {
                                                            oneHunCount++;
                                                            courseFlag = 1;
                                                            break;
                                                        }
                                                    }
                                                }
                                             
                                                //判断标志 如果是1 不进行处理 如果是0 则进行处理
                                                if (courseFlag == 0)
                                                {
                                                    oneWomenCount++;
                                                    courseFlag = 1;

                                                }
                                            }
                                        }
                                        else//老生 也就是 大二、大三、大四 不知道为什么老师没写大四 因此大一作为条件比较好
                                        {
                                            tmpText = tmpText + "      " + this.courseCodeDs.Tables[0].Rows[k][3];
                                            //进行大一的统计
                                            int courseFlag = 0;//匹配标志  
                                            if (tmpCouse.Contains("男"))
                                            {
                                                menCount++;
                                                courseFlag = 1;

                                            }
                                            //[废弃]不是男生
                                            //2016-11-17 还需要匹配一次textbox里面的男生班 没有带男字标志的
                                            else
                                            {
                                                //没有男字标志的男生班

                                                for (int t = 0; t < manClassStrArray.Length; t++)
                                                {
                                                    if (manClassStrArray[t] == tmpCouse)
                                                    {
                                                        menCount++;
                                                        courseFlag = 1;
                                                        break;
                                                    }
                                                }
                                                //混合
                                                if (courseFlag == 0)
                                                {
                                                    for (int t = 0; t < hunClassStrArray.Length; t++)
                                                    {
                                                        if (hunClassStrArray[t] == tmpCouse)
                                                        {
                                                            hunCount++;
                                                            courseFlag = 1;
                                                            break;
                                                        }
                                                    }
                                                }

                                                //判断标志 如果是1 不进行处理 如果是0 则进行处理
                                                if (courseFlag == 0)
                                                {
                                                    womenCount++;
                                                    courseFlag = 1;

                                                }
                                            }
                                        }
                                        //MessageBox.Show(tmpText);
                                        //修改新的单元格内容
                                        curentCell.Value = tmpText;
                                        curentCell.ColumnWidth = 25;
                                        // MessageBox.Show(curentCell.Text.ToString());
                                        courseCount++;
                                        //输出状态
                                        statusLabel.Text = "正在处理" + statusText + " 第 " + courseCount.ToString() + " 条数据...";

                                        break;
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                            }
                        }
                        //这里是每个校区的处理结束
                        //进行内容的标记
                        statusCountText += statusText + "共(" + courseCount + ") :"+ System.Environment.NewLine;
                        statusCountText += "新生：男(" + oneMenCount + ")女(" + oneWomenCount + ")混(" + oneHunCount+ ") " +System.Environment.NewLine;
                        statusCountText += "老生：男(" + menCount + ")女(" + womenCount + ")混(" + hunCount+ ") " +System.Environment.NewLine;
                        statusLabel.Text = statusText + " 排课处理完成";
                    }
                    //保存文件
                    // _wbk.Save();
                    //退出
                    courseApp.AlertBeforeOverwriting = false;
                    _wbk.Close(null, null, null);
                    wbks.Close();
                    courseApp.Quit();
                    //释放掉多余的excel进程
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(courseApp);
                    courseApp = null;
                    ////输出状态
                    statusCountTextBox.Text = statusCountText;
                    statusLabel.Text = "排课处理完成...Powered by postbird";

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "发生错误");
                }
                finally
                {
                    //恢复按钮的可用性
                    courseFileButton.Enabled = true;
                    courseCodeFileButton.Enabled = true;
                    resetButton.Enabled = true;
                    submitButton.Enabled = true;
                }
            }
        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string link = "http://www.ptbird.cn";
            System.Diagnostics.Process.Start(link);
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string link = "http://contact.ptbird.cn";
            System.Diagnostics.Process.Start(link);
        }

        private void statusLabel_Click(object sender, EventArgs e)
        {

        }
    }
}
