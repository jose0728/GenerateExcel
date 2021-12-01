using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace 生成Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook;

            var curPath = AppDomain.CurrentDomain.BaseDirectory;
            using (FileStream stream = new FileStream(curPath + "\\Export\\ExcelDemo.xls", FileMode.Open, FileAccess.Read))
            {
                workbook = new HSSFWorkbook(stream);
            }

            ISheet sheet1 = workbook.GetSheetAt(0);
            IRow row0 = sheet1.GetRow(0);
            ICell cell0 = row0.GetCell(0);

            var oldValue = cell0.StringCellValue;
            var newValue = oldValue.Replace("startTime", DateTime.Parse("2021/05/17 12:00").ToString("yyyyMMdd")).Replace("endTime", DateTime.Parse("2021/05/18 12:00").ToString("yyyyMMdd"));
            cell0.SetCellValue(newValue);

            var list = new List<Model>();
            for (var i = 0; i < 100; i++)
            {
                list.Add(new Model { No = i.ToString().PadLeft(10, '0'), Name = "樊" + i, Time = DateTime.Now.AddSeconds(i).ToString("yyyy-MM-dd HH:mm:ss") });
            }

            ExcelCellInfoModel[] cellInfos = ExcelCellInfoModel.Build(new string[] { "No", "Name", "Time" });
            cellInfos[0].CellType = 1;

            if (list.Count > 0)
            {
                for (int i = 0; i < list.Count; i++)
                {
                    IRow row = sheet1.CreateRow(i + 2);
                    for (int j = 0; j < cellInfos.Length; j++)
                    {
                        ICell cell = row.CreateCell(j);
                        cell.SetCellValue(GetDeviceCellValue(list[i], cellInfos[j].CellEnName));
                    }
                    //最后一行超出
                    row.CreateCell(cellInfos.Length).SetCellValue(" ");
                }
            }


            //转为字节数组
            MemoryStream streamOut = new MemoryStream();
            workbook.Write(streamOut);
            streamOut.Seek(0, SeekOrigin.Begin);

            var res = streamOut.ToArray();

            FileStream stream1 = new FileStream(curPath + $"\\Files\\ExcelDemo{DateTime.Now.ToString("yyyyMMddHHmmss")}.xls", FileMode.CreateNew, FileAccess.Write);
            stream1.Write(res);
            stream1.Close();
        }

        public static string GetDeviceCellValue(Model itemInfo, string cellEnName)
        {
            string result = "";
            switch (cellEnName)
            {
                case "No":
                    result = itemInfo.No;
                    break;
                case "Name":
                    result = itemInfo.Name;
                    break;
                case "Time":
                    result = itemInfo.Time;
                    break;
                default:
                    result = "";
                    break;
            }
            return result;
        }
    }

    public class Model
    {
        public string No { get; set; }
        public string Name { get; set; }
        public string Time { get; set; }
    }

    public class ExcelCellInfoModel
    {
        /// <summary>
        /// 数据列英文名称
        /// </summary>
        public string CellEnName { get; set; }

        /// <summary>
        /// 1文本 2数字
        /// </summary>
        public int CellType { get; set; }

        /// <summary>
        /// 使用名称生成一组，类型为数字的
        /// </summary>
        /// <param name="names"></param>
        /// <returns></returns>
        public static ExcelCellInfoModel[] Build(string[] names)
        {
            ExcelCellInfoModel[] ret = new ExcelCellInfoModel[names.Length];
            for (int i = 0; i < names.Length; i++)
            {
                ret[i] = new ExcelCellInfoModel { CellEnName = names[i], CellType = 2 };
            }
            return ret;
        }
    }
}
