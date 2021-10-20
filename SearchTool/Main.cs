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
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SearchTool
{
    public partial class Main : Form
    {
        private List<DataTable> dataTableList = new List<DataTable>();

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
            LoadSettings();
        }

        /// <summary>
        /// 加载配置
        /// </summary>
        private void LoadSettings()
        {
            try
            {
                var index = 1;
                var loadText = $"{index}.正在检测数据文件...";
                index++;
                richTextBox1.Text = loadText;
                var root = Application.StartupPath;
                if (File.Exists($"{root}\\TestQuestion.xls") || File.Exists($"{root}\\TestQuestion.xlsx"))
                {
                    loadText += $"{Environment.NewLine}{index}.正在配置数据连接驱动通道...";
                    index++;
                    richTextBox1.Text = loadText;
                    using (FileStream fileStream = new FileStream($"{root}\\TestQuestion.xls", FileMode.Open))
                    {
                        dataTableList = new List<DataTable>();
                        dataTableList = ExcelOperationHelper.ExcelStreamToDataTable(fileStream);
                    }
                    if (dataTableList == null || !dataTableList.Any())
                    {
                        loadText += $"{Environment.NewLine}Excel数据文件“TestQuestion”暂无数据!";
                        richTextBox1.Text = loadText;
                        return;
                    }

                    for (int i = 0; i < dataTableList.Count; i++)
                    {
                        loadText += $"{Environment.NewLine}{index}.加载{dataTableList[i].TableName}...";
                        index++;
                        richTextBox1.Text = loadText;
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

                    loadText += $"{Environment.NewLine}{index}.初始化完成。{Environment.NewLine}请输入您要查询的信息，可按空格分隔，将以两个词并联搜索";
                    richTextBox1.Text = loadText;
                }
                else
                {
                    loadText += $"{Environment.NewLine}Excel数据文件“TestQuestion”不存在!";
                    richTextBox1.Text = loadText;
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
            //LoadSettings();
            //textBox1.Text = string.Empty;
            //richTextBox1.Text += $"{Environment.NewLine}重载配置完成。";
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
            Search();
        }

        /// <summary>
        /// 根据关键字搜索 可按空格分隔，将以两个词搜索
        /// </summary>
        private void Search()
        {
            // 重置文本
            richTextBox1.Text = string.Empty;
            // 当前搜索关键字
            var keys = textBox1.Text.Split(" ");
            // 当前选中tab索引
            var curSelectTabIndex = tabControl1.SelectedIndex;
            if (dataTableList != null && dataTableList.Any())
            {
                var curTabDataTable = dataTableList[curSelectTabIndex];
                if (curTabDataTable != null)
                {
                    var datas = ModelConvertHelper<ExcelModel>.ConvertToModel(curTabDataTable);
                    if (datas != null && datas.Any())
                    {
                        var newKeys = new List<string>();
                        foreach (var key in keys)
                        {
                            if (string.IsNullOrEmpty(key)) continue;
                            newKeys.Add(key);
                        }
                        if (newKeys != null && newKeys.Any())
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
                                var resposeHtml = string.Empty;
                                var splitStr = new string[0];
                                foreach (var item in newDatas)
                                {
                                    splitStr = item.item.Split("|");
                                    resposeHtml += $"{(string.IsNullOrEmpty(resposeHtml) ? "" : Environment.NewLine)}【id:{item.id}】 - 【type:{item.type}】{Environment.NewLine}";
                                    foreach (var sp in splitStr)
                                    {
                                        if (string.IsNullOrEmpty(sp)) continue;
                                        resposeHtml += $"{sp}{Environment.NewLine}";
                                    }
                                }
                                richTextBox1.Text = resposeHtml;
                                foreach (var newKey in newKeys)
                                {
                                    if (string.IsNullOrEmpty(newKey)) continue;
                                    ChangeKeyColor(newKey, Color.Red);
                                }
                            }
                        }
                    }
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
            Search();
        }

        /// <summary>
        /// 查重
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            // 重置文本
            richTextBox1.Text = string.Empty;
            Message msg = new Message();
            msg.Show();
            // 当前选中tab索引
            var curSelectTabIndex = tabControl1.SelectedIndex;
            if (dataTableList != null && dataTableList.Any())
            {
                var curTabDataTable = dataTableList[curSelectTabIndex];
                if (curTabDataTable != null)
                {
                    var datas = ModelConvertHelper<ExcelModel>.ConvertToModel(curTabDataTable);
                    if (datas != null && datas.Any())
                    {
                        IAnalyser analyser = new SimHashAnalyser();
                        var likeness = 0.0;
                        var resposeHtml = string.Empty;
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
                                    resposeHtml += $"=========相似度{likeness * 100}%=========";
                                    splitStr1 = excelModel1.item.Split("|");
                                    resposeHtml += $"{(string.IsNullOrEmpty(resposeHtml) ? "" : Environment.NewLine)}【id:{excelModel1.id}】 - 【type:{excelModel1.type}】{Environment.NewLine}";
                                    foreach (var sp in splitStr1)
                                    {
                                        if (string.IsNullOrEmpty(sp)) continue;
                                        resposeHtml += $"{sp}{Environment.NewLine}";
                                    }

                                    splitStr2 = excelModel2.item.Split("|");
                                    resposeHtml += $"{(string.IsNullOrEmpty(resposeHtml) ? "" : Environment.NewLine)}【id:{excelModel2.id}】 - 【type:{excelModel2.type}】{Environment.NewLine}";
                                    foreach (var sp in splitStr2)
                                    {
                                        if (string.IsNullOrEmpty(sp)) continue;
                                        resposeHtml += $"{sp}{Environment.NewLine}";
                                    }
                                    resposeHtml += $"============================{Environment.NewLine}{Environment.NewLine}";
                                }
                            }
                        }
                        msg.Close();
                        if (!string.IsNullOrEmpty(resposeHtml))
                        {
                            richTextBox1.Text = resposeHtml;
                            newKeys = newKeys.Distinct().ToList();
                            if (newKeys != null && newKeys.Any())
                            {
                                foreach (var newKey in newKeys)
                                {
                                    if (string.IsNullOrEmpty(newKey)) continue;
                                    ChangeKeyColor(newKey, Color.Red);
                                }
                            }
                            MessageBox.Show("查重完成！");
                        }
                        else
                            MessageBox.Show("查重完成，不存在相似度大于等于90%的数据！");
                    }
                }
            }
        }
    }
}