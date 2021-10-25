using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SearchTool
{
    public partial class Main : Form
    {
        private List<DataTable> dataTableList = new List<DataTable>();
        private readonly string _fileName = "TestQuestion";
        private int pageIndex = 1;
        private int pageSize = 10;
        private int pages = 1;
        private int curRichTextDataNum = 1;
        private List<string> newKeys1 = new List<string>();// 搜索字符串关键字集合
        private List<string> richTextBoxList1 = new List<string>();// 搜索字符串分页数据集合
        private List<string> newKeys2 = new List<string>();// 查重字符串关键字集合
        private List<string> richTextBoxList2 = new List<string>();// 查重字符串分页数据集合
        private int curSelectTabIndex = 0;// 当前选中tab索引

        /// <summary>
        /// 主函数
        /// </summary>
        public Main()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 主页面加载
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Main_Load(object sender, EventArgs e)
        {
            textBox2.LostFocus += TextBox2_LostFocus;
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
            curRichTextDataNum = 1;
            // 重置搜索字符串关键字集合
            newKeys1 = new List<string>();
            // 重置搜索字符串分页数据集合
            richTextBoxList1 = new List<string>();
            // 重置查重字符串关键字集合
            newKeys2 = new List<string>();
            // 重置查重字符串分页数据集合
            richTextBoxList2 = new List<string>();
        }

        /// <summary>
        /// 总页数赋值
        /// </summary>
        /// <param name="pages"></param>
        private void InitPages(int pages)
        {
            textBox3.Text = pages.ToString();
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

                            newDatas = newDatas.Skip((pageIndex - 1) * pageSize).Take(pageSize);
                            var resposeHtml = string.Empty;
                            var splitStr = new string[0];
                            foreach (var item in newDatas)
                            {
                                splitStr = item.item.Split("|");
                                richTextBox1.AppendText($"{(string.IsNullOrEmpty(richTextBox1.Text) ? "" : Environment.NewLine)}{item.type}{Environment.NewLine}");
                                foreach (var sp in splitStr)
                                {
                                    if (string.IsNullOrEmpty(sp)) continue;
                                    richTextBox1.AppendText($"{sp}{Environment.NewLine}");
                                }
                                curRichTextDataNum++;
                            }
                            richTextBoxList1.Add(richTextBox1.Text);
                            ChangeKeyColor1();
                            if (pages > 1)
                            {
                                //using (BackgroundWorker bw = new BackgroundWorker())
                                //{
                                //    //后台线程需要执行的委托
                                //    bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Thread1);
                                //    //后台线程结束后 会调用该委托
                                //    //bw.DoWork += new DoWorkEventHandler(Thread1);
                                //    //如果线程需要参数，可以传入参数  DoWorkEventArgs e.Argument调用参数
                                //    bw.RunWorkerAsync();
                                //}
                                Thread th = new Thread(new ThreadStart(Thread1));
                                th.IsBackground = true;
                                th.Start();
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
                                curRichTextDataNum = pageSize + 1;
                                var resposeHtml = string.Empty;
                                var splitStr = new string[0];
                                foreach (var item in newDatas)
                                {
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
                                    curRichTextDataNum++;
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 搜索关键字重新改变颜色
        /// </summary>
        private void ChangeKeyColor1()
        {
            if (newKeys1 != null && newKeys1.Any())
            {
                foreach (var newKey in newKeys1)
                {
                    if (string.IsNullOrEmpty(newKey)) continue;
                    ChangeKeyColor(newKey, Color.Red);
                }
            }
        }

        /// <summary>
        /// 查重关键字重新改变颜色
        /// </summary>
        private void ChangeKeyColor2()
        {
            if (newKeys2 != null && newKeys2.Any())
            {
                foreach (var newKey in newKeys2)
                {
                    if (string.IsNullOrEmpty(newKey)) continue;
                    ChangeKeyColor(newKey, Color.Red);
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
        /// 当前选项卡选中改变事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
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
            RestartSettings();
            Message msg = new Message();
            msg.Show();
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
                        IAnalyser analyser = new SimHashAnalyser();
                        var likeness = 0.0;
                        var excelModel1 = new ExcelModel();
                        var excelModel2 = new ExcelModel();
                        var splitStr1 = new string[0];
                        var splitStr2 = new string[0];
                        var newKeys = new List<string>();
                        for (int i = 0; i < datas.Count; i++)
                        {
                            for (int j = i + 1; j < datas.Count; j++)
                            {
                                excelModel1 = datas[i];
                                excelModel2 = datas[j];
                                likeness = analyser.GetLikenessValue(excelModel1.item, excelModel2.item);
                                if (likeness >= 0.9)
                                {
                                    newKeys.Add($"相似度{likeness * 100}%");
                                    richTextBox1.AppendText($"=========相似度{likeness * 100}%=========");
                                    splitStr1 = excelModel1.item.Split("|");
                                    richTextBox1.AppendText($"{(string.IsNullOrEmpty(richTextBox1.Text) ? "" : Environment.NewLine)}{excelModel1.type}{Environment.NewLine}");
                                    foreach (var sp in splitStr1)
                                    {
                                        if (string.IsNullOrEmpty(sp)) continue;
                                        richTextBox1.AppendText($"{sp}{Environment.NewLine}");
                                    }
                                    splitStr2 = excelModel2.item.Split("|");
                                    richTextBox1.AppendText($"{(string.IsNullOrEmpty(richTextBox1.Text) ? "" : Environment.NewLine)}{excelModel2.type}{Environment.NewLine}");
                                    foreach (var sp in splitStr2)
                                    {
                                        if (string.IsNullOrEmpty(sp)) continue;
                                        richTextBox1.AppendText($"{sp}{Environment.NewLine}");
                                    }
                                    richTextBox1.AppendText($"============================{Environment.NewLine}{Environment.NewLine}");
                                    if (curRichTextDataNum == pageSize)
                                    {
                                        curRichTextDataNum++;
                                        break;
                                    }
                                    curRichTextDataNum++;
                                }
                                if (curRichTextDataNum == pageSize + 1) break;
                            }
                        }
                        msg.Close();
                        if (!string.IsNullOrEmpty(richTextBox1.Text))
                        {
                            newKeys = newKeys.Distinct().ToList();
                            newKeys2 = newKeys;
                            richTextBoxList2.Add(richTextBox1.Text);
                            ChangeKeyColor2();
                            using (BackgroundWorker bw = new BackgroundWorker())
                            {
                                //后台线程需要执行的委托
                                bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Thread2);
                                //后台线程结束后 会调用该委托
                                //bw.DoWork += new DoWorkEventHandler(Thread2);
                                //如果线程需要参数，可以传入参数  DoWorkEventArgs e.Argument调用参数
                                bw.RunWorkerAsync();
                            }
                            MessageBox.Show($"“{curTabDataTable.TableName}”查重完成！");
                        }
                        else
                            MessageBox.Show($"查重完成，“{curTabDataTable.TableName}”不存在相似度大于等于90%的数据！");
                    }
                    else
                    {
                        msg.Close();
                        MessageBox.Show($"“{curTabDataTable.TableName}”暂无数据或数据格式不正确！");
                    }
                }
            }
        }

        /// <summary>
        /// 查重线程
        /// </summary>
        private void Thread2(object sender, RunWorkerCompletedEventArgs e)
        {
            if (dataTableList != null && dataTableList.Any())
            {
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
                        var newKeys = new List<string>();
                        var resposeHtml = string.Empty;
                        curRichTextDataNum = 1;
                        var lastAddIndex = 1;
                        for (int i = 0; i < datas.Count; i++)
                        {
                            for (int j = i + 1; j < datas.Count; j++)
                            {
                                excelModel1 = datas[i];
                                excelModel2 = datas[j];
                                likeness = analyser.GetLikenessValue(excelModel1.item, excelModel2.item);
                                if (likeness >= 0.9)
                                {
                                    newKeys.Add($"相似度{likeness * 100}%");
                                    resposeHtml += $"=========相似度{likeness * 100}%=========";
                                    splitStr1 = excelModel1.item.Split("|");
                                    resposeHtml += $"{Environment.NewLine}{excelModel1.type}{Environment.NewLine}";
                                    foreach (var sp in splitStr1)
                                    {
                                        if (string.IsNullOrEmpty(sp)) continue;
                                        resposeHtml += $"{sp}{Environment.NewLine}";
                                    }
                                    splitStr2 = excelModel2.item.Split("|");
                                    resposeHtml += $"{Environment.NewLine}{excelModel2.type}{Environment.NewLine}";
                                    foreach (var sp in splitStr2)
                                    {
                                        if (string.IsNullOrEmpty(sp)) continue;
                                        resposeHtml += $"{sp}{Environment.NewLine}";
                                    }
                                    resposeHtml += $"============================{Environment.NewLine}{Environment.NewLine}";
                                    if (curRichTextDataNum % pageSize == 0)
                                    {
                                        if (!string.IsNullOrEmpty(resposeHtml) && curRichTextDataNum != pageSize)
                                        {
                                            richTextBoxList2.Add(resposeHtml);
                                            lastAddIndex = curRichTextDataNum;
                                        }
                                        resposeHtml = String.Empty;
                                    }
                                    curRichTextDataNum++;
                                }
                            }
                        }
                        if (!string.IsNullOrEmpty(resposeHtml) && (curRichTextDataNum > lastAddIndex || lastAddIndex == 1))
                            richTextBoxList2.Add(resposeHtml);
                        newKeys = newKeys.Distinct().ToList();
                        newKeys2 = newKeys;
                        pages = richTextBoxList2.Count();
                        this.textBox3.Text = pages.ToString();
                    }
                }
            }
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
            if (richTextBoxList1.Count() > 0 && richTextBoxList2.Count() == 0)
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
            if (richTextBoxList2.Count() > 0 && richTextBoxList1.Count() == 0)
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
            if (richTextBoxList1.Count() > 0 && richTextBoxList2.Count() == 0)
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
            if (richTextBoxList2.Count() > 0 && richTextBoxList1.Count() == 0)
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
            if (richTextBoxList1.Count() > 0 && richTextBoxList2.Count() == 0)
            {
                // 搜索
                var curPageIndex = Convert.ToInt32(textBox2.Text);
                pageIndex = curPageIndex;
                var richText = richTextBoxList1.Skip(pageIndex - 1).Take(1).FirstOrDefault() ?? String.Empty;
                richTextBox1.Text = richText;
                ChangeKeyColor1();
                textBox2.Text = pageIndex.ToString();
            }
            if (richTextBoxList2.Count() > 0 && richTextBoxList1.Count() == 0)
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
    }
}