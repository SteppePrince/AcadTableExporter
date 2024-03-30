using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System;
using System.Linq;

namespace AcadTableExporter
{
    public class AcadTableExporter
    {
        [CommandMethod("ExportTableToCSV")]
        public void ExportTableToCSVCommand()
        {

            Document acDoc = Application.DocumentManager.MdiActiveDocument;

            Database acCurDb = acDoc.Database;

            Editor acEd = acDoc.Editor;

            PromptSelectionResult selectionResult = acEd.GetSelection();

            if (selectionResult.Status != PromptStatus.OK) return;

            SelectionSet selectionSet = selectionResult.Value;
            
            List<Line> lines = new List<Line>();
            List<DBText> texts = new List<DBText>();

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {

                foreach (SelectedObject selObj in selectionSet)
                {

                    if (selObj != null)
                    {

                        Entity entity = acTrans.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;

                        if (entity is Line line)
                        {
                            lines.Add(line);
                        }
                        else if (entity is DBText text)
                        {
                            texts.Add(text);
                        }
                    }
                }

                // 输出直线坐标用于调试
                foreach (var line in lines)
                {
                    acEd.WriteMessage($"\nLine Start: ({line.StartPoint.X}, {line.StartPoint.Y}), End: ({line.EndPoint.X}, {line.EndPoint.Y})");
                }

                // 分析表格结构并生成CSV内容
                string csvContent = GenerateCSVContent();

                // 尝试获取LOCALROOTPREFIX系统变量值
                string localRootPrefix = Application.GetSystemVariable("LOCALROOTPREFIX") as string ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                // 组合最终的文件路径
                string filePath = Path.Combine(localRootPrefix, "ExportedTable.csv");
                // 用于拼接目录路径和文件名，创建完整的文件路径
                File.WriteAllText(filePath, csvContent, Encoding.UTF8);
                acEd.WriteMessage($"\nCSV file created at: {filePath}");
                acTrans.Commit();
            }

            string GenerateCSVContent()
            {
                // 初始化CSV构建器
                StringBuilder csvBuilder = new StringBuilder();

                // 提取唯一的X坐标并排序
                var xCoords = lines.SelectMany(line => new[] { Math.Round(line.StartPoint.X, 3), Math.Round(line.EndPoint.X, 3) })
                   .Distinct()
                   .OrderBy(x => x)
                   .ToList();

                // 提取唯一的Y坐标并排序
                var yCoords = lines.SelectMany(line => new[] { Math.Round(line.StartPoint.Y, 3), Math.Round(line.EndPoint.Y, 3) })
                   .Distinct()
                   .OrderByDescending(y => y)
                   .ToList();

                // 输出X坐标和Y坐标用于调试
                acEd.WriteMessage("\nX Coordinates:");
                foreach (var x in xCoords)
                {
                    acEd.WriteMessage($" {x}");
                }

                acEd.WriteMessage("\nY Coordinates:");
                foreach (var y in yCoords)
                {
                    acEd.WriteMessage($" {y}");
                }

                // 表格数据的二维数组
                string[,] tableData = new string[yCoords.Count - 1, xCoords.Count - 1];

                // 定位每个文本到对应的单元格
                foreach (var text in texts)
                {
                    // 查找文本对应的列索引
                    var colIndex = xCoords.FindIndex(x => x > text.Position.X) - 1;
                    // 查找文本对应的行索引
                    var rowIndex = yCoords.FindIndex(y => y < text.Position.Y) - 1;

                    // 调试输出文本位置和计算出的索引
                    acEd.WriteMessage($"\nText '{text.TextString}' at ({text.Position.X}, {text.Position.Y}) mapped to cell [{rowIndex}, {colIndex}]");

                    // 如果索引有效，将文本的字符串内容(TextString)放入之前定义的tableData二维数组的相应位置
                    // tableData数组的每个元素代表表格的一个单元格，行和列索引对应于数组的维度
                    if (rowIndex >= 0 && colIndex >= 0)
                    {
                        tableData[rowIndex, colIndex] = text.TextString;
                    }
                }

                // 构建CSV内容
                for (int i = 0; i < tableData.GetLength(0); i++)
                {
                    var rowData = new List<string>();
                    for (int j = 0; j < tableData.GetLength(1); j++)
                    {
                        rowData.Add(tableData[i, j] ?? "");
                    }
                    csvBuilder.AppendLine(string.Join(",", rowData));
                }

                // 输出CSV构建的完整内容用于调试
                acEd.WriteMessage($"\nCSV Content:\n{csvBuilder.ToString()}");

                return csvBuilder.ToString();

            }

        }



    }
}
