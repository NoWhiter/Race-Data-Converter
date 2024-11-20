using NPOI.SS.Formula.Functions;
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

        // 添加一个新的字段来存储队伍信息
        private Dictionary<string, string> playerTeamMap = new Dictionary<string, string>();

        // 添加新的字段来存储队伍名单文件路径
        private string teamListFilePath;

        int numError = 0;

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Text = "选择文件";
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                // 选择对话框标题
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

        private void button4_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "选择队伍名单文件";
                openFileDialog.Multiselect = false;
                openFileDialog.Filter = "csv文件|*.csv|所有文件|*.*";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    teamListFilePath = openFileDialog.FileName;
                    try
                    {
                        // 读取队伍名单文件
                        using (StreamReader reader = new StreamReader(teamListFilePath))
                        {
                            playerTeamMap.Clear(); // 清除之前的队伍信息
                            string line;
                            int lineCount = 0;
                            while ((line = reader.ReadLine()) != null)
                            {
                                string[] fields = line.Split(',');
                                if (fields.Length >= 2)
                                {
                                    // 使用玩家名字作为键，存储队伍名称
                                    playerTeamMap[fields[0].Trim()] = fields[1].Trim();
                                    lineCount++;
                                }
                            }
                            ((Button)sender).Text = $"已选择队伍名单：{Path.GetFileName(teamListFilePath)}";
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("读取队伍名单失败：" + ex.Message);
                        teamListFilePath = null;
                        playerTeamMap.Clear();
                    }
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

            // 使用 using 语句确保资源释放
            using (var package = new ExcelPackage())
            {
                if (package.Workbook == null)
                {
                    MessageBox.Show("创建ExcelWorkbook失败！");
                    return;
                }

                // 创建工作表
                // 添加 "总览" Sheet
                ExcelWorksheet overviewSheet = package.Workbook.Worksheets.Add("总览");
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

                        // 写入名
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
                // 创建玩家数据分数表
                overviewSheet = package.Workbook.Worksheets.Add("分数");
                row = 2;
                col = 1;
                List<Player> player_ID = new List<Player>();

                // 遍历所有谱面收集玩家信息
                foreach (var currentBeatmap in beatmapToLines)
                {
                    foreach (string dataLine in currentBeatmap.Value)
                    {
                        string[] fields = dataLine.Split(",");
                        // 检查玩家是否已存在
                        bool playerExists = false;
                        foreach (var player in player_ID)
                        {
                            if (player.Id == fields[2])
                            {
                                playerExists = true;
                                break;
                            }
                        }
                        // 如果玩家不存在，则添加
                        if (!playerExists)
                        {
                            player_ID.Add(new Player(fields[2]));
                        }
                    }
                }

                // 写入玩家ID和名称
                foreach (var player in player_ID)
                {
                    // 在所有谱面中查找该玩家的名称
                    foreach (var currentBeatmap in beatmapToLines)
                    {
                        foreach (string dataLine in currentBeatmap.Value)
                        {
                            string[] fields = dataLine.Split(",");
                            if (fields[2] == player.Id)
                            {
                                overviewSheet.Cells[row, 1].Value = fields[2]; // ID
                                overviewSheet.Cells[row, 2].Value = fields[3]; // 名称
                                break;
                            }
                        }
                        if (overviewSheet.Cells[row, 1].Value != null)
                        {
                            break;
                        }
                    }
                    row++;
                }

                // 写入队伍信息
                row = 2;
                foreach (var player in player_ID)
                {
                    foreach (var currentBeatmap in beatmapToLines)
                    {
                        foreach (string dataLine in currentBeatmap.Value)
                        {
                            string[] fields = dataLine.Split(",");
                            if (fields[2] == player.Id)
                            {
                                overviewSheet.Cells[row, 3].Value = fields[0];
                                break;
                            }
                        }
                        if (overviewSheet.Cells[row, 3].Value != null)
                        {
                            break;
                        }
                    }
                    row++;
                }

                // 写入每个谱面的分数
                row = 2;
                col = 4;
                matchline = 4;
                foreach (var currentBeatmap in beatmapToLines)
                {
                    overviewSheet.Cells[1, matchline].Value = currentBeatmap.Key;
                    foreach (var player in player_ID)
                    {
                        bool scoreFound = false;
                        foreach (string dataLine in currentBeatmap.Value)
                        {
                            string[] fields = dataLine.Split(",");
                            if (fields[2] == player.Id)
                            {
                                overviewSheet.Cells[row, col].Value = fields[4];
                                scoreFound = true;
                                break;
                            }
                        }
                        if (!scoreFound)
                        {
                            // 如果玩家没有这张图的成绩，可以选择留空或填写特定值
                            overviewSheet.Cells[row, col].Value = null;
                        }
                        row++;
                    }
                    row = 2;
                    matchline += 2;
                    col += 2;
                }

                if (playerTeamMap.Count > 0)
                {
                    // 添加"队伍"表格
                    overviewSheet = package.Workbook.Worksheets.Add("队伍");
                    row = 2;
                    matchline = 1;

                    // 用于存储每张图的最后写入行
                    Dictionary<string, int> lastRowByMap = new Dictionary<string, int>();

                    foreach (var currentBeatmap in beatmapToLines)
                    {
                        string mapId = currentBeatmap.Key;
                        bool isNewMap = !lastRowByMap.ContainsKey(mapId);

                        // 如果是新的图号，写入图号并初始化起始行
                        if (isNewMap)
                        {
                            // 写入图号
                            overviewSheet.Cells[1, matchline].Value = mapId;
                            lastRowByMap[mapId] = row;
                        }

                        // 用于存储当前图的队伍信息
                        Dictionary<string, List<string[]>> teamPlayers = new Dictionary<string, List<string[]>>();

                        // 收集该谱面下的玩家信息并按队伍分组
                        foreach (string dataLine in currentBeatmap.Value)
                        {
                            string[] fields = dataLine.Split(',');
                            if (fields.Length >= 8)
                            {
                                string playerName = fields[3];  // 玩家名字

                                // 如果找到了该选手的队伍信息
                                string teamName;
                                if (playerTeamMap.TryGetValue(playerName, out teamName))
                                {
                                    if (!teamPlayers.ContainsKey(teamName))
                                    {
                                        teamPlayers[teamName] = new List<string[]>();
                                    }
                                    teamPlayers[teamName].Add(fields);
                                }
                            }
                        }

                        // 获取当前图的起始列
                        int currentMatchline = matchline;

                        // 获取当前图的起始行
                        int currentRow = lastRowByMap[mapId];

                        // 当前遍历到第几个
                        int currentPlayerIndex = 0;

                        // 临时分数计算变量
                        int score1 = 0;
                        int score2 = 0;

                        // 记录上一个队伍名称
                        string previousTeamName = null;

                        // 写入每个队伍的信息
                        foreach (var teamEntry in teamPlayers)
                        {
                            string teamName = teamEntry.Key;
                            var teamData = teamEntry.Value;

                            // 如果队伍名称变化了且不是第一个队伍，添加空行
                            if (previousTeamName != null && teamName != previousTeamName)
                            {
                                currentRow++;
                            }

                            // 记录当前组的起始行
                            int groupStartRow = currentRow;
                            int playerCount = 0;

                            // 写入该队伍的选手信息
                            foreach (var playerData in teamData)
                            {
                                // 写入选手信息 (fields[3]-fields[7])
                                for (int i = 3; i <= 7; i++)
                                {
                                    // 写入队伍名
                                    overviewSheet.Cells[currentRow, currentMatchline].Value = teamName;
                                    overviewSheet.Cells[currentRow, currentMatchline + i - 2].Value = playerData[i];
                                }
                                currentRow++;
                                playerCount++;
                            }

                            // 如果有选手数据，计算总分
                            if (playerCount > 0)
                            {
                                // 计算这组所有选手的总分
                                int totalScore = 0;
                                for (int i = 0; i < playerCount; i++)
                                {
                                    var score = overviewSheet.Cells[groupStartRow + i, currentMatchline + 2].Value;
                                    if (score != null)
                                    {
                                        totalScore += Convert.ToInt32(score);
                                    }
                                }
                                
                                // 为每个队员写入相同的总分
                                for (int i = 0; i < playerCount; i++)
                                {
                                    overviewSheet.Cells[groupStartRow + i, currentMatchline + 6].Value = totalScore;
                                }
                            }

                            // 更新上一个队伍名称
                            previousTeamName = teamName;
                        }

                        // 更新该图的最后写入行
                        lastRowByMap[mapId] = currentRow;

                        // 如果是新的图号，移动到下一个图的位置
                        if (isNewMap)
                        {
                            matchline += 7;
                        }

                        // 在处理完所有队伍后，每张图进行排序
                        // 添加在 if (isNewMap) { matchline += 7; } 之前

                        // 获取当前图的所有数据范围
                        int startRow = lastRowByMap[mapId] - currentRow + 2; // 第一个队伍开始的行
                        int endRow = currentRow - 1; // 最后一个队伍结束的行

                        // 创建临时表存储排序后的数据
                        var tempSheet = package.Workbook.Worksheets.Add("Temp");

                        // 收集所有行的数据并排序
                        var rowData = new List<(int row, object[] values, int totalScore)>();
                        for (int dataRow = startRow; dataRow <= endRow; dataRow++)
                        {
                            if (overviewSheet.Cells[dataRow, currentMatchline].Value != null) // 跳过空行
                            {
                                var values = new object[7];
                                for (int Tempcol = 0; Tempcol <= 6; Tempcol++)
                                {
                                    values[Tempcol] = overviewSheet.Cells[dataRow, currentMatchline + Tempcol].Value;
                                }
                                int totalScore = Convert.ToInt32(values[6]); // 总分在第7列
                                rowData.Add((dataRow, values, totalScore));
                            }
                        }

                        // 按总分降序排序
                        var sortedData = rowData.OrderByDescending(x => x.totalScore).ToList();

                        // 将排序后的数据写回临时表
                        int newRow = startRow;
                        for (int i = 0; i < sortedData.Count; i++)
                        {
                            var data = sortedData[i];
                            for (int Newcol = 0; Newcol <= 6; Newcol++)
                            {
                                tempSheet.Cells[newRow, currentMatchline + Newcol].Value = data.values[Newcol];
                            }
                            newRow++;
                        }

                        // 将排序后的数据复制回原表
                        for (int copyRow = startRow; copyRow <= endRow; copyRow++)
                        {
                            for (int copyCol = currentMatchline; copyCol <= currentMatchline + 6; copyCol++)
                            {
                                overviewSheet.Cells[copyRow, copyCol].Value = tempSheet.Cells[copyRow, copyCol].Value;
                            }
                        }

                        // 合并相同分数的单元格
                        int mergeStartRow = startRow;
                        for (int i = 1; i <= sortedData.Count; i++)
                        {
                            if (i == sortedData.Count || sortedData[i].totalScore != sortedData[i - 1].totalScore)
                            {
                                // 如果分数不同或是最后一行，且有多于一行需要合并
                                if (mergeStartRow < startRow + i - 1)
                                {
                                    var cell1 = overviewSheet.Cells[mergeStartRow, currentMatchline + 6, startRow + i - 1, currentMatchline + 6];
                                    cell1.Merge = true;
                                    var cell2 = overviewSheet.Cells[mergeStartRow, currentMatchline, startRow + i - 1, currentMatchline];
                                    cell2.Merge = true;
                                }
                                mergeStartRow = startRow + i;
                            }
                        }

                        // 删除临时表
                        package.Workbook.Worksheets.Delete("Temp");
                    }
                }

                // 设置所有工作表的所有单元格居中
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    // 获取工作表使用的范围
                    var dimension = worksheet.Dimension;
                    if (dimension != null)
                    {
                        // 选择整个使用范围
                        var range = worksheet.Cells[dimension.Address];
                        
                        // 设置水平和垂直居中
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        range.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    }
                }

                // 保存文件
                try
                {
                    // 检查文件是否被占用
                    if (File.Exists(excelOutputPath))
                    {
                        File.Delete(excelOutputPath);
                    }

                    // 使用 FileStream 来确保文件被正确覆盖
                    using (var stream = new FileStream(excelOutputPath, FileMode.Create, FileAccess.Write))
                    {
                        package.SaveAs(stream);
                        numError = 0;
                    }
                }
                catch (IOException ex)
                {
                    MessageBox.Show("请关闭以前生成的同名文件！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    numError = 1;
                }
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

            // 处理 scorev2Data 中的每个谱面的每个家的数据
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

            bool hasError = false; // 添加一个标志来跟踪是否发生错误

            foreach (string csvFilePath in selectedFilePaths)
            {
                string excelOutputPath = GetExcelOutputPath(csvFilePath);

                // 检查文件名是否与正在打开的文件名相同
                if (selectedFilePaths.Contains(excelOutputPath))
                {
                    MessageBox.Show("处理文件时发生错误，请检查文件是否被占用！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    hasError = true;
                    continue;
                }

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
                    hasError = true;
                }
            }

            if (hasError == false && numError != 1) // 只有在没有错误时才显示完成窗口
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

        private void label2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Thanks Kuit and WOSHIZHAZHA120 for help", "鸣谢", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}