using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SearchTool
{
    public partial class Main : Form
    {
        SynchronizationContext m_SyncContext = null;
        private List<DataTable> dataTableList = new List<DataTable>();
        private readonly string _fileName = "TestQuestion";
        private int pageIndex = 1;
        private int pageSize = 20;
        private int pages = 1;
        private int curRichTextDataNum = 0;
        private List<string> newKeys1 = new List<string>();// 搜索字符串关键字集合
        private List<string> richTextBoxList1 = new List<string>();// 搜索字符串分页数据集合
        private List<string> newKeys2 = new List<string>();// 查重字符串关键字集合
        private List<string> richTextBoxList2 = new List<string>();// 查重字符串分页数据集合
        private int curSelectTabIndex = 0;// 当前选中tab索引
        private bool isSearch = true;// 给分页功能相关按钮用 默认是 搜索
        private double ProgressLabTxtLoadingIndex = 0;// 后台查重进度百分比
        private int proIndex = 0;// 查重循环匹配次数
        private int TotalNum = 0;// 查重循环总匹配次数
        private Thread th1;// 搜索线程
        private Thread th2;// 查重线程

        /// <summary>
        /// 主函数
        /// </summary>
        public Main()
        {
            InitializeComponent();
            //获取UI线程同步上下文
            m_SyncContext = SynchronizationContext.Current;
        }

        /// <summary>
        /// 主页面加载
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Main_Load(object sender, EventArgs e)
        {
            textBox2.LostFocus += TextBox2_LostFocus;
            label3.Parent = progressBar1;
            label3.Location = new Point(650, 3);
            LoadSettings();
        }

        /// <summary>
        /// 当前页失去焦点事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <exception cref="NotImplementedException"></exception>
        private void TextBox2_LostFocus(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox2.Text))
            {
                textBox2.Text = "1";
            }
        }

        /// <summary>
        /// 重置配置参数
        /// </summary>
        private void RestartSettings()
        {
            // 销毁线程
            if (th1 != null && th1.IsAlive)
            {
                th1_run.Text = "0";
            }
            if (th2 != null && th2.IsAlive)
            {
                th2_run.Text = "0";
            }
            // 重置富文本内容
            richTextBox1.Text = String.Empty;
            // 重置当前页
            pageIndex = 1;
            // 重置当前页文本框
            textBox2.Text = "1";
            // 重置总页数
            pages = 1;
            // 重置总页数文本框
            textBox3.Text = "1";
            // 重置富文本内容条数
            curRichTextDataNum = 0;
            // 重置搜索字符串关键字集合
            newKeys1 = new List<string>();
            // 重置搜索字符串分页数据集合
            richTextBoxList1 = new List<string>();
            // 重置查重字符串关键字集合
            newKeys2 = new List<string>();
            // 重置查重字符串分页数据集合
            richTextBoxList2 = new List<string>();
            // 重置查重进度条
            progressBar1.Value = 0;
            // 重置查重匹配次数
            proIndex = 0;
            // 查重循环总匹配次数
            TotalNum = 0;
            // 重置查重进度百分比数值
            ProgressLabTxtLoadingIndex = 0;
            // 重置后台查重进度文本
            label3.Text = "";
            // 重置搜索或查重结果总条数
            richNum.Text = "搜索或查重结果总条数：0";
        }

        /// <summary>
        /// 总页数赋值
        /// </summary>
        /// <param name="pages"></param>
        private void InitPages(object pages)
        {
            textBox3.Text = pages.ToString();
            textBox3.Refresh();
        }

        /// <summary>
        /// 搜索结果总条数赋值
        /// </summary>
        /// <param name="pages"></param>
        private void InitPageTotalNum1(object num)
        {
            richNum.Text = $"搜索或查重结果总条数：{num}";
            richNum.Refresh();
        }
        
        /// <summary>
        /// 查重结果总条数赋值
        /// </summary>
        /// <param name="pages"></param>
        private void InitPageTotalNum2(object num)
        {
            richNum.Text = $"搜索或查重结果总条数：{curRichTextDataNum}";
            richNum.Refresh();
        }

        /// <summary>
        /// 加载配置
        /// </summary>
        private void LoadSettings()
        {
            try
            {
                RestartSettings();
                var index = 1;
                richTextBox1.AppendText($"{index}.正在检测数据文件...");
                index++;
                var root = Application.StartupPath;
                var path = string.Empty;
                if (File.Exists($"{root}..\\{_fileName}.xls"))
                {
                    path = $"{root}..\\{_fileName}.xls";
                }
                if (File.Exists($"{root}..\\{_fileName}.xlsx"))
                {
                    path = $"{root}..\\{_fileName}.xlsx";
                }
                if (!string.IsNullOrEmpty(path))
                {
                    richTextBox1.AppendText($"{Environment.NewLine}{index}.正在配置数据连接驱动通道...");
                    index++;
                    using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        dataTableList = new List<DataTable>();
                        dataTableList = ExcelOperationHelper.ExcelStreamToDataTable(fileStream);
                    }
                    if (dataTableList == null || !dataTableList.Any())
                    {
                        richTextBox1.AppendText($"{Environment.NewLine}Excel数据文件“{_fileName}”暂无数据!");
                        return;
                    }
                    for (int i = 0; i < dataTableList.Count; i++)
                    {
                        richTextBox1.AppendText($"{Environment.NewLine}{index}.加载{dataTableList[i].TableName}...");
                        index++;
                        if (i == 0)
                        {
                            tabPage1.Text = dataTableList[i].TableName;
                        }
                        else
                        {
                            TabPage tp = new TabPage($"tablePage{i + 1}");
                            tp.Text = dataTableList[i].TableName;
                            tabControl1.TabPages.Add(tp);
                        }
                    }
                    richTextBox1.AppendText($"{Environment.NewLine}{index}.加载图片识别文字功能...");
                    index++;
                    var secret = EstateCertOCR.InitSecret();
                    if (secret == null || string.IsNullOrEmpty(secret.secretId) || string.IsNullOrEmpty(secret.secretKey))
                    {
                        richTextBox1.AppendText("图别识别密钥加载失败,功能暂不可用,请检查secret_Id_Key.txt文件是否存在!");
                    }
                    else
                        richTextBox1.AppendText("加载成功!");
                    richTextBox1.AppendText($"{Environment.NewLine}{index}.初始化完成。{Environment.NewLine}请输入您要查询的信息，可按空格分隔，将以两个词并联搜索");
                }
                else
                {
                    richTextBox1.AppendText($"{Environment.NewLine}Excel数据文件“TestQuestion”不存在!");
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }
        }

        /// <summary>
        /// 重载配置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            Process process = Process.GetCurrentProcess();
            process.Close();
            Application.Restart();
        }

        /// <summary>
        /// 搜索框文本改变事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            isSearch = true;
            RestartSettings();
            Search();
        }

        /// <summary>
        /// 根据关键字搜索 可按空格分隔，将以两个词搜索
        /// </summary>
        private void Search()
        {
            // 当前搜索关键字
            var keys = textBox1.Text.Split(" ");
            var newKeys = new List<string>();
            foreach (var key in keys)
            {
                if (string.IsNullOrEmpty(key)) continue;
                newKeys.Add(key);
            }
            if (newKeys == null || !newKeys.Any()) return;
            newKeys1 = newKeys;
            // 当前选中tab索引
            curSelectTabIndex = tabControl1.SelectedIndex;
            if (dataTableList != null && dataTableList.Any())
            {
                var curTabDataTable = dataTableList[curSelectTabIndex];
                if (curTabDataTable != null)
                {
                    var datas = ModelConvertHelper<ExcelModel>.ConvertToModel(curTabDataTable).Where(_ => _.type != null && _.item != null).ToList();
                    if (datas != null && datas.Any())
                    {
                        var searchResults = new List<string>();
                        var expression = PredicateExtensions.True<ExcelModel>();
                        foreach (var newKey in newKeys)
                        {
                            expression = expression.And(_ => _.item.Contains(newKey));
                        }
                        var predicate = expression.Compile();
                        var newDatas = datas.Where(predicate);
                        if (newDatas != null && newDatas.Any())
                        {
                            // 更新总页数
                            pages = newDatas.Count() % pageSize == 0 ? newDatas.Count() / pageSize : newDatas.Count() / pageSize + 1;
                            InitPages(pages);
                            InitPageTotalNum1(newDatas.Count());
                            newDatas = newDatas.Skip((pageIndex - 1) * pageSize).Take(pageSize);
                            var resposeHtml = string.Empty;
                            var splitStr = new string[0];
                            foreach (var item in newDatas)
                            {
                                curRichTextDataNum++;
                                Application.DoEvents();
                                splitStr = item.item.Split("|");
                                richTextBox1.AppendText($"{(string.IsNullOrEmpty(richTextBox1.Text) ? "" : Environment.NewLine)}{item.type}{Environment.NewLine}");
                                foreach (var sp in splitStr)
                                {
                                    if (string.IsNullOrEmpty(sp)) continue;
                                    richTextBox1.AppendText($"{sp}{Environment.NewLine}");
                                }
                            }
                            richTextBoxList1.Add(richTextBox1.Text);
                            ChangeKeyColor1();
                            if (pages > 1)
                            {
                                this.th1_run.Text = "1";
                                th1 = new Thread(new ThreadStart(Thread1));
                                th1.IsBackground = true;
                                th1.Start();
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 搜索线程
        /// </summary>
        private void Thread1()
        {
            if (newKeys1 != null && newKeys1.Any())
            {
                if (dataTableList != null && dataTableList.Any())
                {
                    this.Invoke(new Action(() =>
                    {
                        curSelectTabIndex = tabControl1.SelectedIndex;
                    }));
                    var curTabDataTable = dataTableList[curSelectTabIndex];
                    if (curTabDataTable != null)
                    {
                        var datas = ModelConvertHelper<ExcelModel>.ConvertToModel(curTabDataTable).Where(_ => _.type != null && _.item != null).ToList();
                        if (datas != null && datas.Any())
                        {
                            var searchResults = new List<string>();
                            var expression = PredicateExtensions.True<ExcelModel>();
                            foreach (var newKey in newKeys1)
                            {
                                expression = expression.And(_ => _.item.Contains(newKey));
                            }
                            var predicate = expression.Compile();
                            var newDatas = datas.Where(predicate);
                            if (newDatas != null && newDatas.Any())
                            {
                                newDatas = newDatas.Skip(pageSize);
                                var resposeHtml = string.Empty;
                                var splitStr = new string[0];
                                foreach (var item in newDatas)
                                {
                                    curRichTextDataNum++;
                                    Application.DoEvents();
                                    this.Invoke(new Action(() =>
                                    {
                                        if (this.th1_run.Text.Equals("0"))
                                            return;
                                    }));
                                    splitStr = item.item.Split("|");
                                    resposeHtml += $"{(curRichTextDataNum % pageSize == 1 ? "" : Environment.NewLine)}{item.type}{Environment.NewLine}";
                                    foreach (var sp in splitStr)
                                    {
                                        if (string.IsNullOrEmpty(sp)) continue;
                                        resposeHtml += $"{sp}{Environment.NewLine}";
                                    }
                                    if (curRichTextDataNum % pageSize == 0 || curRichTextDataNum == pageSize + newDatas.Count())
                                    {
                                        richTextBoxList1.Add(resposeHtml);
                                        resposeHtml = String.Empty;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            this.th1_run.Text = "1";
        }

        /// <summary>
        /// 搜索关键字重新改变颜色
        /// </summary>
        /// <param name="isMainUIThread">是否UI主线程调用,默认是</param>
        private void ChangeKeyColor1(bool isMainUIThread = true)
        {
            if (newKeys1 != null && newKeys1.Any())
            {
                foreach (var newKey in newKeys1)
                {
                    if (string.IsNullOrEmpty(newKey)) continue;
                    if (isMainUIThread)
                        ChangeKeyColor(newKey, Color.Red);
                    else
                        ChangeKeyColorThread(newKey, Color.Red);
                }
            }
        }

        /// <summary>
        /// 查重关键字重新改变颜色
        /// </summary>
        /// <param name="isMainUIThread">是否UI主线程调用,默认是</param>
        private void ChangeKeyColor2(bool isMainUIThread = true)
        {
            if (newKeys2 != null && newKeys2.Any())
            {
                foreach (var newKey in newKeys2.ToList())
                {
                    if (string.IsNullOrEmpty(newKey)) continue;
                    if (isMainUIThread)
                        ChangeKeyColor(newKey, Color.Red);
                    else
                        ChangeKeyColorThread(newKey, Color.Red);
                }
            }
        }

        /// <summary>
        /// 切换字体颜色
        /// </summary>
        /// <param name="key">关键字</param>
        /// <param name="color">颜色</param>
        private void ChangeKeyColor(string key, Color color)
        {
            Regex regex = new Regex(key);
            //找出内容中所有的要替换的关键字
            MatchCollection collection = regex.Matches(richTextBox1.Text);
            //对所有的要替换颜色的关键字逐个替换颜色
            foreach (Match match in collection)
            {
                //开始位置、长度、颜色缺一不可
                richTextBox1.SelectionStart = match.Index;
                richTextBox1.SelectionLength = key.Length;
                richTextBox1.SelectionColor = color;
            }
        }

        /// <summary>
        /// 切换字体颜色 子线程专用
        /// </summary>
        /// <param name="key"></param>
        /// <param name="color"></param>
        private void ChangeKeyColorThread(string key, Color color)
        {
            this.Invoke(new Action(() =>
            {
                Regex regex = new Regex(key);
                //找出内容中所有的要替换的关键字
                MatchCollection collection = regex.Matches(richTextBox1.Text);
                //对所有的要替换颜色的关键字逐个替换颜色
                foreach (Match match in collection)
                {
                    //开始位置、长度、颜色缺一不可
                    richTextBox1.SelectionStart = match.Index;
                    richTextBox1.SelectionLength = key.Length;
                    richTextBox1.SelectionColor = color;
                }
            }));
        }

        /// <summary>
        /// 当前选项卡选中改变事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            isSearch = true;
            RestartSettings();
            Search();
        }

        /// <summary>
        /// 查重
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            isSearch = false;
            RestartSettings();
            this.th2_run.Text = "1";
            th2 = new Thread(new ThreadStart(Thread2));
            th2.IsBackground = true;
            th2.Start();
        }

        /// <summary>
        /// 查重子线程
        /// </summary>
        private void Thread2()
        {
            if (dataTableList != null && dataTableList.Any())
            {
                this.Invoke((Action)(() =>
                {
                    curSelectTabIndex = tabControl1.SelectedIndex;
                }));
                var curTabDataTable = dataTableList[curSelectTabIndex];
                if (curTabDataTable != null)
                {
                    var datas = ModelConvertHelper<ExcelModel>.ConvertToModel(curTabDataTable).Where(_ => _.type != null && _.item != null).ToList();
                    if (datas != null && datas.Any())
                    {
                        IAnalyser analyser = new SimHashAnalyser();
                        var likeness = 0.0;
                        var excelModel1 = new ExcelModel();
                        var excelModel2 = new ExcelModel();
                        var splitStr1 = new string[0];
                        var splitStr2 = new string[0];
                        var resposeHtml = string.Empty;
                        var sigleResposeHtml = string.Empty;
                        var lastAddIndex = 1;
                        TotalNum = GetTotalNum(datas.Count());
                        this.Invoke(new Action(() =>
                        {
                            progressBar1.Maximum = TotalNum;
                        }));
                        for (int i = 0; i < datas.Count; i++)
                        {
                            for (int j = i + 1; j < datas.Count; j++)
                            {
                                this.Invoke(new Action(() =>
                                {
                                    Application.DoEvents();
                                    if (this.th2_run.Text.Equals("0"))
                                    {
                                        datas = new List<ExcelModel>();
                                        return;
                                    }
                                    else
                                    {
                                        proIndex++;
                                        excelModel1 = datas[i];
                                        excelModel2 = datas[j];
                                        likeness = analyser.GetLikenessValue(excelModel1.item, excelModel2.item);
                                        if (likeness >= 0.9)
                                        {
                                            curRichTextDataNum++;
                                            newKeys2.Add($"相似度{likeness * 100}%");
                                            newKeys2 = newKeys2.Distinct().ToList();
                                            resposeHtml += $"=========相似度{likeness * 100}%=========";
                                            sigleResposeHtml += $"=========相似度{likeness * 100}%=========";
                                            splitStr1 = excelModel1.item.Split("|");
                                            resposeHtml += $"{Environment.NewLine}{excelModel1.type}{Environment.NewLine}";
                                            sigleResposeHtml += $"{Environment.NewLine}{excelModel1.type}{Environment.NewLine}";
                                            foreach (var sp in splitStr1)
                                            {
                                                if (string.IsNullOrEmpty(sp)) continue;
                                                resposeHtml += $"{sp}{Environment.NewLine}";
                                                sigleResposeHtml += $"{sp}{Environment.NewLine}";
                                            }
                                            splitStr2 = excelModel2.item.Split("|");
                                            resposeHtml += $"{Environment.NewLine}{excelModel2.type}{Environment.NewLine}";
                                            sigleResposeHtml += $"{Environment.NewLine}{excelModel2.type}{Environment.NewLine}";
                                            foreach (var sp in splitStr2)
                                            {
                                                if (string.IsNullOrEmpty(sp)) continue;
                                                resposeHtml += $"{sp}{Environment.NewLine}";
                                                sigleResposeHtml += $"{sp}{Environment.NewLine}";
                                            }
                                            resposeHtml += $"============================{Environment.NewLine}{Environment.NewLine}";
                                            sigleResposeHtml += $"============================{Environment.NewLine}{Environment.NewLine}";
                                            if (curRichTextDataNum <= pageSize)
                                                m_SyncContext.Post(SetRichTextAppendText, sigleResposeHtml);
                                            m_SyncContext.Post(InitPageTotalNum2, curRichTextDataNum);
                                            if (curRichTextDataNum % pageSize == 0)
                                            {
                                                richTextBoxList2.Add(resposeHtml);
                                                pages = richTextBoxList2.Count();
                                                m_SyncContext.Post(SetTextSafePost, pages);
                                                lastAddIndex = curRichTextDataNum;
                                                resposeHtml = String.Empty;
                                            }
                                            sigleResposeHtml = String.Empty;
                                        }
                                        var proText = Math.Round((double)proIndex * 100 / TotalNum, 2);
                                        if (proText > ProgressLabTxtLoadingIndex)
                                            ProgressLabTxtLoadingIndex = proText;
                                        if (proIndex == TotalNum)
                                            ProgressLabTxtLoadingIndex = 100;
                                        m_SyncContext.Post(SetProgressLabTxtLoading, $"{proIndex}/{TotalNum}");
                                        m_SyncContext.Post(SetProgressPerformStep, "0");
                                    }
                                }));
                            }
                        }
                        if (!string.IsNullOrEmpty(resposeHtml) && curRichTextDataNum > lastAddIndex && lastAddIndex != 1)
                            richTextBoxList2.Add(resposeHtml);
                        pages = richTextBoxList2.Count();
                        //this.textBox3.Text = pages.ToString();
                        //在线程中更新UI（通过UI线程同步上下文m_SyncContext）
                        m_SyncContext.Post(SetTextSafePost, pages);
                    }
                }
            }
            this.Invoke(new Action(() =>
            {
                this.th2_run.Text = "1";
            }));
        }

        /// <summary>
        /// 当前页输入事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            //如果输入的不是退格和数字，则屏蔽输入
            if (!(e.KeyChar == '\b' || (e.KeyChar >= '0' && e.KeyChar <= '9')))
            {
                e.Handled = true;
            }
        }

        /// <summary>
        /// 上一页
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            if (richTextBoxList1.Count() == 0 && richTextBoxList2.Count() == 0)
            {
                MessageBox.Show("请先输入要搜索的关键字或点击查重按钮!");
                return;
            }
            if (isSearch)
            {
                if (richTextBoxList1.Count() > 0)
                {
                    // 搜索
                    if (textBox2.Text.Equals("1"))
                    {
                        pageIndex = 1;
                        MessageBox.Show("已经是第一页了!");
                        return;
                    }
                    else
                    {
                        var curPageIndex = Convert.ToInt32(textBox2.Text);
                        pageIndex = curPageIndex - 1;
                        var richText = richTextBoxList1.Skip(pageIndex - 1).Take(1).FirstOrDefault() ?? String.Empty;
                        richTextBox1.Text = richText;
                        ChangeKeyColor1();
                        textBox2.Text = pageIndex.ToString();
                    }
                }
            }
            else
            {
                if (richTextBoxList2.Count() > 0)
                {
                    // 查重
                    if (textBox2.Text.Equals("1"))
                    {
                        pageIndex = 1;
                        MessageBox.Show("已经是第一页了!");
                        return;
                    }
                    else
                    {
                        var curPageIndex = Convert.ToInt32(textBox2.Text);
                        pageIndex = curPageIndex - 1;
                        var richText = richTextBoxList2.Skip(pageIndex - 1).Take(1).FirstOrDefault() ?? String.Empty;
                        richTextBox1.Text = richText;
                        ChangeKeyColor2();
                        textBox2.Text = pageIndex.ToString();
                    }
                }
            }
        }

        /// <summary>
        /// 下一页
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            if (richTextBoxList1.Count() == 0 && richTextBoxList2.Count() == 0)
            {
                MessageBox.Show("请先输入要搜索的关键字或点击查重按钮!");
                return;
            }
            if (isSearch)
            {
                if (richTextBoxList1.Count() > 0)
                {
                    // 搜索
                    if (textBox2.Text.Equals(pages.ToString()))
                    {
                        pageIndex = pages;
                        MessageBox.Show("已经是最后一页了!");
                        return;
                    }
                    else
                    {
                        var curPageIndex = Convert.ToInt32(textBox2.Text);
                        pageIndex = curPageIndex + 1;
                        var richText = richTextBoxList1.Skip(pageIndex - 1).Take(1).FirstOrDefault() ?? String.Empty;
                        richTextBox1.Text = richText;
                        ChangeKeyColor1();
                        textBox2.Text = pageIndex.ToString();
                    }
                }
            }
            else
            {
                if (richTextBoxList2.Count() > 0)
                {
                    // 查重
                    if (textBox2.Text.Equals(pages.ToString()))
                    {
                        pageIndex = pages;
                        MessageBox.Show("已经是最后一页了!");
                        return;
                    }
                    else
                    {
                        var curPageIndex = Convert.ToInt32(textBox2.Text);
                        pageIndex = curPageIndex + 1;
                        var richText = richTextBoxList2.Skip(pageIndex - 1).Take(1).FirstOrDefault() ?? String.Empty;
                        richTextBox1.Text = richText;
                        ChangeKeyColor2();
                        textBox2.Text = pageIndex.ToString();
                    }
                }
            }
        }

        /// <summary>
        /// 跳转到
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            if (richTextBoxList1.Count() == 0 && richTextBoxList2.Count() == 0)
            {
                MessageBox.Show("请先输入要搜索的关键字或点击查重按钮!");
                return;
            }
            if (isSearch)
            {
                if (richTextBoxList1.Count() > 0)
                {
                    // 搜索
                    var curPageIndex = Convert.ToInt32(textBox2.Text);
                    pageIndex = curPageIndex;
                    var richText = richTextBoxList1.Skip(pageIndex - 1).Take(1).FirstOrDefault() ?? String.Empty;
                    richTextBox1.Text = richText;
                    ChangeKeyColor1();
                    textBox2.Text = pageIndex.ToString();
                }
            }
            else
            {
                if (richTextBoxList2.Count() > 0)
                {
                    // 查重
                    var curPageIndex = Convert.ToInt32(textBox2.Text);
                    pageIndex = curPageIndex;
                    var richText = richTextBoxList2.Skip(pageIndex - 1).Take(1).FirstOrDefault() ?? String.Empty;
                    richTextBox1.Text = richText;
                    ChangeKeyColor2();
                    textBox2.Text = pageIndex.ToString();
                }
            }
        }

        /// <summary>
        /// 当前页输入值改变事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox2.Text))
            {
                var curPageIndex = Convert.ToInt32(textBox2.Text);
                var curPages = Convert.ToInt32(textBox3.Text);
                if (curPageIndex < 1 || curPageIndex > curPages)
                {
                    MessageBox.Show($"当前页范围：{1}-{pages}");
                    textBox2.Text = "1";
                    return;
                }
            }
        }

        /// <summary>
        /// 鼠标拖动某项到搜索文本框工作区时发生事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }

        /// <summary>
        /// 拖动到搜索文本框工作区操作完成时发生事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // 从拖动数据里得到路径
                string realpath = ((Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
                if (string.IsNullOrEmpty(realpath))
                {
                    MessageBox.Show("图片路径为空,请重新拖动图片");
                    return;
                }
                // 文件后缀名
                string p = Path.GetExtension(realpath);
                if (string.IsNullOrEmpty(p))
                {
                    MessageBox.Show("暂时只支持识别jgp/jpeg/png/bmp格式的图片");
                    return;
                }
                if (p == ".jpg" || p == ".jpeg" || p == ".png" || p == ".bmp")
                {
                    var ocrStr = EstateCertOCR.Ocr(realpath);
                    if (!string.IsNullOrEmpty(ocrStr))
                    {
                        textBox1.Text = ocrStr;
                    }
                    else
                    {
                        MessageBox.Show("图片中未检测到文本");
                    }
                }
                else
                {
                    MessageBox.Show("暂时只支持识别jgp/jpeg/png/bmp格式的图片");
                }
            }
            else if (e.Data.GetDataPresent(DataFormats.Text))
            {
                var txt = (string)e.Data.GetData(DataFormats.Text);
                if (!string.IsNullOrEmpty(txt))
                {
                    textBox1.Text = txt;
                }
            }
            else if (e.Data.GetDataPresent(DataFormats.Rtf))
            {
                var rtf = (string)e.Data.GetData(DataFormats.Rtf);
                if (!string.IsNullOrEmpty(rtf))
                {
                    richTextBox1.Rtf = rtf;
                    richTextBox1.Copy();
                    richTextBox1.Rtf = String.Empty;
                    if (Clipboard.ContainsImage())
                    {
                        var imgBase64 = EstateCertOCR.ImgToBase64(new Bitmap(Clipboard.GetImage()));
                        var ocrStr = EstateCertOCR.Ocr("", true, imgBase64);
                        if (!string.IsNullOrEmpty(ocrStr))
                        {
                            textBox1.Text = ocrStr;
                        }
                        else
                        {
                            MessageBox.Show("图片中未检测到文本");
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("请拖动或粘贴要识别的文字或图片到搜索框");
            }
        }

        /// <summary>
        /// 计算查重需要的总数量
        /// </summary>
        /// <param name="datasCount"></param>
        /// <returns></returns>
        private int GetTotalNum(int datasCount)
        {
            var r = 0;
            if (datasCount > 2)
            {
                for (int i = 1; i < datasCount; i++)
                {
                    r += datasCount - i;
                }
            }
            return r;
        }

        /// <summary>
        /// 跨线程更新富文本
        /// </summary>
        /// <param name="text"></param>
        private void SetRichTextAppendText(object text)
        {
            this.richTextBox1.AppendText(text.ToString());
            this.richTextBox1.Refresh();
            ChangeKeyColor2(false);
            this.richTextBox1.Refresh();
        }

        /// <summary>
        /// 跨线程更新总页数
        /// </summary>
        /// <param name="text"></param>
        private void SetTextSafePost(object text)
        {
            this.textBox3.Text = text.ToString();
            this.textBox3.Refresh();
        }

        /// <summary>
        /// 跨线程更新后台查重进度文本
        /// </summary>
        private void SetProgressLabTxtLoading(object text)
        {
            this.label3.Text = $"后台查重进度{ProgressLabTxtLoadingIndex}%（{text}）";
            this.label3.Refresh();
        }

        /// <summary>
        /// 跨线程更新进度条
        /// </summary>
        /// <param name="text"></param>
        private void SetProgressPerformStep(object text)
        {
            this.progressBar1.PerformStep();
        }

        /// <summary>
        /// 搜索输入框按下按键并且释放后事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //[Ctrl+V]
            if (e.KeyChar == 22)
            {
                if (Clipboard.ContainsImage())
                {
                    var imgBase64 = EstateCertOCR.ImgToBase64(new Bitmap(Clipboard.GetImage()));
                    var ocrStr = EstateCertOCR.Ocr("", true, imgBase64);
                    if (!string.IsNullOrEmpty(ocrStr))
                    {
                        textBox1.Text = ocrStr;
                    }
                    else
                    {
                        MessageBox.Show("图片中未检测到文本");
                    }
                }
                if (Clipboard.ContainsFileDropList())
                {
                    // 从拖动数据里得到路径
                    string realpath = ((Array)Clipboard.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
                    if (string.IsNullOrEmpty(realpath))
                    {
                        MessageBox.Show("图片路径为空,请重新复制图片");
                        return;
                    }
                    // 文件后缀名
                    string p = Path.GetExtension(realpath);
                    if (string.IsNullOrEmpty(p))
                    {
                        MessageBox.Show("暂时只支持识别jgp/jpeg/png/bmp格式的图片");
                        return;
                    }
                    if (p == ".jpg" || p == ".jpeg" || p == ".png" || p == ".bmp")
                    {
                        var ocrStr = EstateCertOCR.Ocr(realpath);
                        if (!string.IsNullOrEmpty(ocrStr))
                        {
                            textBox1.Text = ocrStr;
                        }
                        else
                        {
                            MessageBox.Show("图片中未检测到文本");
                        }
                    }
                    else
                    {
                        MessageBox.Show("暂时只支持识别jgp/jpeg/png/bmp格式的图片");
                    }
                }
            }
        }
    }
}