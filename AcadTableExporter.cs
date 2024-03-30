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
            // 获取当前活动的AutoCAD文档（即用户当前正在操作的DWG文件）
            // Application.DocumentManager是管理所有打开文档的文档管理器，而MdiActiveDocument属性返回最顶层或当前激活的文档
            // 在AutoCAD中，可以同时打开多个文档，但MdiActiveDocument只引用当前用户正直接操作的那一个
            Document acDoc = Application.DocumentManager.MdiActiveDocument;

            // 每个AutoCAD文档都有一个Database对象，其中存储了该文档的所有几何和非几何信息，比如图形、线条、文字、图层等
            // 这行代码通过acDoc.Database访问了当前文档的数据库，允许你读取和修改这些数据。数据库是存储所有图形数据的容器
            Database acCurDb = acDoc.Database;

            // Editor对象是与用户交互的接口，比如选择对象、获取用户输入和在命令行显示消息等
            // 这行代码通过acDoc.Editor获取了当前文档的编辑器实例，使得插件可以执行这些交互操作
            Editor acEd = acDoc.Editor;

            // GetSelection() 方法调用弹出一个选择框，让用户可以在图纸上通过绘制矩形框来选择对象
            // 用户可以通过点击和拖拽鼠标来确定这个矩形框的位置和大小，从而选择框内的所有对象
            // 这个方法返回一个 PromptSelectionResult 对象，它包含了用户选择的结果，包括用户是否成功进行了选择，以及具体选择了哪些对象
            PromptSelectionResult selectionResult = acEd.GetSelection();

            // 这行代码检查用户是否成功完成了选择
            // 如果用户没有成功完成选择（例如，用户取消了选择操作），则 Status 属性不会等于 PromptStatus.OK
            // 在这种情况下，代码通过 return 语句直接退出当前的命令或方法
            // 这是一种简单的错误处理方式，确保后续的代码只有在用户成功进行选择后才会执行
            if (selectionResult.Status != PromptStatus.OK) return;

            // 这是用户通过选择框选取的所有图形对象的集合
            // SelectionSet 对象提供了对这些选择对象的访问
            SelectionSet selectionSet = selectionResult.Value;

            // 这两行代码分别初始化了两个列表：lines 和 texts
            // lines 用于存储所有选取的直线对象(Line)，而 texts 用于存储所有选取的文本对象(DBText)
            // Line 和 DBText 都是继承自 Entity 类的AutoCAD图形对象类型，分别代表图纸中的直线和文本
            List<Line> lines = new List<Line>();
            List<DBText> texts = new List<DBText>();

            // Transaction 对象用于管理对数据库的更改，确保这些更改要么完全应用要么完全不应用，以保持数据的一致性
            // 这里通过调用 acCurDb.TransactionManager.StartTransaction() 开启一个新的事务
            // using 语句确保在代码块执行完毕后自动调用 Transaction 对象的 Dispose 方法，这通常会提交或回滚事务，取决于事务中是否发生了异常
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // 遍历选择集合中的所有对象
                // SelectedObject 是 SelectionSet 中的项，代表一个被选择的对象
                // 每个 SelectedObject 包含了对应图形对象的引用信息，如对象的ID(ObjectId)
                foreach (SelectedObject selObj in selectionSet)
                {
                    // 对于每个选择的对象，代码首先检查它是否非空
                    if (selObj != null)
                    {
                        // 使用事务对象(acTrans)的 GetObject 方法和对象的ID(selObj.ObjectId)来获取实际的图形对象
                        // 这里是以只读模式打开(OpenMode.ForRead)
                        // GetObject 方法返回的是一个基类型 DBObject 的对象，这里通过 as Entity 转换为 Entity 类型，因为图形对象在AutoCAD中都是 Entity 类型或其子类的实例
                        Entity entity = acTrans.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
                        // 使用 is 关键字检查对象的具体类型
                        // 如果对象是 Line 类型的直线，就将其添加到 lines 列表中
                        // 如果是 DBText 类型的文本对象，就将其添加到 texts 列表中
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
