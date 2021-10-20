using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
            LoadSettings();
            textBox1.Text = String.Empty;
            richTextBox1.Text += $"{Environment.NewLine}重载配置完成。";
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
            richTextBox1.Text = String.Empty;
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
                                    resposeHtml += $"{(string.IsNullOrEmpty(resposeHtml) ? "" : Environment.NewLine)}{item.id}{item.type}{Environment.NewLine}";
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
            richTextBox1.Text = String.Empty;
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
                        const string HayStack = "中国香港………………";
                        const string Needle = "中国香港 2013………………";

                        IAnalyser analyser = new SimHashAnalyser();
                        var likeness = analyser.GetLikenessValue(Needle, HayStack);
                        MessageBox.Show($"Likeness: {likeness * 100}% ");
                    }
                }
            }
        }
    }
}