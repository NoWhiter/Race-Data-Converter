using OfficeOpenXml;
using System.Diagnostics;

namespace GBC
{
    public partial class Form1 : Form
    {
        // 存储当前选择的文件名
        private string fileName;

        public Form1()
        {
            InitializeComponent();
            // 设置窗口标题
            this.Text = "GBC数据生成器";
        }

        // 导出路径
        private string ExportFilePath { get; set; }
        // 选择的文件路径
        private string selectedFilePath { get; set; }
        // 选择的文件路径列表
        private List<string> selectedFilePaths = new List<string>();

        private void Button1_Click(object sender, EventArgs e)
        {
            button1.Text = "选择文件";
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                // 设置文件选择对话框标题
                openFileDialog.Title = "选择文件";

                // 允许多选文件
                openFileDialog.Multiselect = true;

                // 设置文件筛选器
                openFileDialog.Filter = "csv文件|*.csv|所有文件|*.*";

                // 如果用户选择了文件并点击了确定
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // 保存选择的文件路径
                    selectedFilePaths = openFileDialog.FileNames.ToList();

                    // 更新按钮文本显示选择的文件
                    ((Button)sender).Text = $"已选择文件：{string.Join(", ", selectedFilePaths.Select(Path.GetFileName))}";
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button2.Text = "选择导出";
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                // 设置文件夹选择对话框描述
                folderBrowserDialog.Description = "选择导出目录";

                // 如果用户选择了文件夹并点击了确定
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    // 保存选择的导出路径
                    ExportFilePath = folderBrowserDialog.SelectedPath;

                    // 更新按钮文本显示导出路径
                    ((Button)sender).Text = $"导出目录：{ExportFilePath}";
                }
            }
        }

        private void sortLine(List<String> lines)
        {
            lines.Sort((line1, line2) =>
            {
                string[] fields1 = line1.Split(',');
                string[] fields2 = line2.Split(',');

                if (fields1.Length > 4 && fields2.Length > 4)
                {
                    if (double.TryParse(fields1[4], out double score1) && double.TryParse(fields2[4], out double score2))
                    {
                        // 按分数降序排序
                        return score2.CompareTo(score1);
                    }
                }

                return 0; // 如果无法比较则保持原顺序
            });
        }

        private void outputFile(String excelOutputPath, Dictionary<string, List<string>> beatmapToLines)
        {
            // 设置 EPPlus 许可
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                if (package.Workbook == null)
                {
                    MessageBox.Show("创建ExcelWorkbook失败！");
                    return;
                }

                // 创建工作表
                // 添加 "详细" Sheet
                ExcelWorksheet overviewSheet = package.Workbook.Worksheets.Add("详细");
                int row = 2; //行
                int col = 1; //列
                int matchline = 1;
                foreach (var currentBeatmap in beatmapToLines)
                {
                    overviewSheet.Cells[1, matchline].Value = currentBeatmap.Key;
                    foreach (var dataLine in currentBeatmap.Value)
                    {
                        string[] fields = dataLine.Split(',');

                        // 写入数据字段
                        for (int i = 2; i <= 7; i++)
                        {
                            overviewSheet.Cells[row, col].Value = fields[i];
                            col++;
                        }

                        // 移动到下一行
                        row++;
                        col = matchline;
                    }
                    matchline += 6;
                    col = matchline;
                    row = 2;
                }
                // 添加 "排名" Sheet
                overviewSheet = package.Workbook.Worksheets.Add("排名");
                matchline = 1;
                row = 2;
                col = matchline;
                int rank = 1;
                foreach (var currentBeatmap in beatmapToLines)
                {
                    overviewSheet.Cells[1, matchline].Value = currentBeatmap.Key;
                    foreach (string dataLine in currentBeatmap.Value)
                    {
                        string[] fields = dataLine.Split(',');

                        // 写入基础数据
                        for (int i = 2; i <= 4; i++)
                        {
                            overviewSheet.Cells[row, col].Value = fields[i];
                            col++;
                        }

                        // 写入排名
                        overviewSheet.Cells[row, col].Value = rank;

                        // 移动位置
                        row++;
                        rank++;
                        col = matchline;
                    }
                    rank = 1;
                    matchline += 4;
                    col = matchline;
                    row = 2;
                }
                // 创建玩家数据汇总表
                overviewSheet = package.Workbook.Worksheets.Add("汇总");
                row = 2;
                col = 1;
                List<Player> player_ID = new List<Player>();
                int player_number = 0;
                foreach (var currentBeatmap in beatmapToLines)
                {
                    foreach (string dataLine in currentBeatmap.Value)
                    {
                        string[] fields = dataLine.Split(",");

                        // 写入玩家数据
                        for (int i = 2; i <= 3; i++)
                        {
                            overviewSheet.Cells[row, col].Value = fields[i];
                            col++;
                        }
                        // 添加玩家对象
                        player_ID.Add(new Player(fields[2]));
                        player_number++;
                        // 移动位置
                        row++;
                        col = 1;
                    }
                    break;
                }
                row = 2;
                foreach (var currentBeatmap in beatmapToLines)
                {
                    foreach (string dataLine in currentBeatmap.Value)
                    {
                        string[] fields = dataLine.Split(",");
                        overviewSheet.Cells[row, 3].Value = fields[0];
                        row++;
                    }
                    break;
                }
                row = 2;
                col = 4;
                matchline = 4;
                foreach (var currentBeatmap in beatmapToLines)
                {
                    overviewSheet.Cells[1, matchline].Value = currentBeatmap.Key;
                    foreach (string dataLine in currentBeatmap.Value)
                    {
                        string[] fields = dataLine.Split(",");
                        for(int j = 0;j < player_number; j++)
                        {
                            if (fields[2] == player_ID[j].Id)
                            {
                                overviewSheet.Cells[j + 2, col].Value = fields[4];
                            }
                        }
                    }
                    matchline += 2;
                    col += 2;
                }
                package.SaveAs(new FileInfo(excelOutputPath));
            }
        }

        private void processFile(StreamReader reader, String excelOutputPath)
        {
            // 用于存储所有 "scorev2" 相关的数据
            Dictionary<string, Dictionary<string, List<string>>> scorev2Data = new Dictionary<string, Dictionary<string, List<string>>>();

            string line;
            string currentKey = null;
            string Roomname = null;
            int RoomID;
            while ((line = reader.ReadLine()) != null)
            {
                // 分割每行数据
                string[] fields = line.Split(',');
                // 检查是否为 "scorev2" 相关数据
                if (fields.Length > 3 && fields[3] == "scorev2")
                {
                    double number1;
                    int if_deleted = 0;
                    if (!double.TryParse(fields[5], out number1))
                    {
                        // 处理无效数据
                        currentKey = if_deleted.ToString();
                        if_deleted++;
                    }
                    else
                    {
                        // 使用谱面ID作为键
                        currentKey = fields[8];
                    }
                    if (!scorev2Data.ContainsKey(currentKey))
                    {
                        scorev2Data[currentKey] = new Dictionary<string, List<string>>();
                    }
                }
                else if (fields.Length == 0)
                {
                    continue;
                }
                else if (fields.Length <= 5 && fields.Length > 1)
                {
                    int.TryParse(fields[2], out RoomID);
                    Roomname = fields[3];
                    continue;
                }
                else if (currentKey != null)
                {
                    if (fields.Length == 0 || (fields[0] != "red" && fields[0] != "blue" && fields[0] != "none"))
                    {
                        // 跳过无效数据
                        continue;
                    }

                    string currentName = fields[1];

                    // 确保该玩家在当前谱面下有对应的数据列表
                    if (!scorev2Data[currentKey].ContainsKey(currentName))
                    {
                        scorev2Data[currentKey][currentName] = new List<string>();
                    }

                    // 添加房间名到数据行
                    string lineWithRoomname = Roomname + "," + line;

                    // 添加数据
                    scorev2Data[currentKey][currentName].Add(lineWithRoomname);
                }
            }

            // 处理 scorev2Data 中的每个谱面的每个玩家的数据
            foreach (var currentBeatmap in scorev2Data)
            {
                foreach (var currentUser in currentBeatmap.Value)
                {
                    // 对每个玩家的成绩进行排序
                    sortLine(currentUser.Value);
                }
                if (radioButton1.Checked)
                {
                    // 去重处理
                    // 只保留每个玩家的最高分
                    foreach (var currentUser in currentBeatmap.Value)
                    {
                        if (currentUser.Value.Count > 1)
                        {
                            currentUser.Value.RemoveRange(1, currentUser.Value.Count - 1);
                        }
                    }
                }
            }

            // 整理 beatmapToLines，合并所有玩家的成绩
            Dictionary<string, List<string>> beatmapToLines = new Dictionary<string, List<string>>();
            foreach (var currentBeatmap in scorev2Data)
            {
                beatmapToLines[currentBeatmap.Key] = new List<string>();
                foreach (var currentUser in currentBeatmap.Value)
                {
                    foreach (var currentLine in currentUser.Value)
                    {
                        beatmapToLines[currentBeatmap.Key].Add(currentLine);
                    }
                }
                sortLine(beatmapToLines[currentBeatmap.Key]);
            }

            outputFile(excelOutputPath, beatmapToLines);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (selectedFilePaths.Count == 0)
            {
                MessageBox.Show("请选择文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (ExportFilePath == null)
            {
                MessageBox.Show("请选择导出路径", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            int if_right = 0;
            foreach (string csvFilePath in selectedFilePaths)
            {
                string excelOutputPath = GetExcelOutputPath(csvFilePath);

                try
                {
                    using (StreamReader reader = new StreamReader(csvFilePath))
                    {
                        processFile(reader, excelOutputPath);
                    }
                }
                catch (IOException ex)
                {
                    MessageBox.Show("处理文件时发生错误，请检查文件是否被占用！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if_right++;
                }
            }
            if (if_right == 0)
            {
                MessageBox.Show("完成", "提示", MessageBoxButtons.OK);
                Process.Start("explorer.exe", ExportFilePath);
            }
        }
        private string GetExcelOutputPath(string csvFilePath)
        {
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(csvFilePath);
            if (radioButton1.Checked)
            {
                return Path.Combine(ExportFilePath, $"{fileNameWithoutExtension} - 去重.xlsx");
            }
            else if (radioButton2.Checked)
            {
                return Path.Combine(ExportFilePath, $"{fileNameWithoutExtension} - 不去重.xlsx");
            }
            MessageBox.Show("请选择去重或不去重", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return fileNameWithoutExtension;
        }
    }
}