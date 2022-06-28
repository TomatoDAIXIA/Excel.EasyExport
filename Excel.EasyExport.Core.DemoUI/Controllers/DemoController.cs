using Microsoft.AspNetCore.Mvc;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.Streaming;
using Excel.EasyExport;
using Excel.EasyExport.Core.DemoUI.Controllers;

namespace Excel.EasyExport.Core.DemoUI.Controllers
{
    [ApiController]
    [Route("[controller]/[action]")]
    public class DemoController : ControllerBase
    {

        private readonly ILogger<DemoController> _logger;

        public DemoController(ILogger<DemoController> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// 基本使用，自动映射字段 和类型
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IActionResult Demo1()
        {
            IWorkbook wb = new XSSFWorkbook();
            //Prop 对应实体的字段名，区分大小写
            List<ExportColumn> columns = new()
            {
                new(){ Label="姓名", Prop="Name"},
                new(){ Label="底薪", Prop="BaseSalary"},
                new(){ Label="是否全勤", Prop="IsFullAttendance",},
                new(){ Label="交通", Prop="Transportation"},
                new(){ Label="餐补", Prop="Meal",},
                new(){ Label="创建时间", Prop=nameof(Salary.CreateTime)},
            };

            List<Salary> data = new List<Salary>();
            data.Add(new Salary()
            {
                Name = "张三",
                BaseSalary = 8000,
                IsFullAttendance = true,
                Meal = 300,
                Transportation = 300,
                CreateTime = DateTime.Now
            });

            var sheet = EasyExport.CreateSheet(wb, columns, data);

            using (MemoryStream ms = new())
            {
                wb.Write(ms);
                wb.Close();
                return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}.xlsx");
            }
        }

        /// <summary>
        /// 合并表头
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IActionResult Demo2()
        {
            IWorkbook wb = new XSSFWorkbook();

            List<ExportColumn> columns = new()
            {
                new(){ Label="姓名", Prop="Name",},
                new()
                {
                    Label = "工资",
                    Children = new() {
                        new(){ Label="底薪", Prop="BaseSalary"},
                        new(){ Label="全勤", Prop="FullAttendance"},
                        new(){ Label="补助",
                            Children = new () {
                                new (){ Label="交通", Prop="Transportation"},
                                new (){ Label="餐补", Prop="Meal"},
                            },
                        },
                    },
                },
                new(){ Label="创建时间",Prop = nameof(Salary.CreateTime)},
            };

            List<Salary> data = new() {
                new() { Name = "张三",
                    BaseSalary = 8000,
                    FullAttendance = 200,
                    Meal = 300,
                    Transportation = 300,
                    CreateTime = DateTime.Now
                }
            };

            EasyExport.CreateSheet(wb, columns, data, "testSheet");

            using (MemoryStream ms = new())
            {
                wb.Write(ms);
                wb.Close();
                return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}.xlsx");
            }
        }

        /// <summary>
        /// 指定列样式
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IActionResult Demo3()
        {
            IWorkbook wb = new XSSFWorkbook();

            List<ExportColumn> columns = new()
            {
                new(){
                    Label="姓名",
                    Prop="Name",
                    Width = 100,
                    Style = ()=>{
                        ICellStyle style = wb.CreateCellStyle();
                        style.Alignment = HorizontalAlignment.Left;
                        IFont font = wb.CreateFont();
                        font.IsBold=true;
                        style.SetFont(font);
                        return style;
                    }
                },
                new(){ Label="底薪", Prop="BaseSalary"},
                new(){ Label="全勤", Prop="FullAttendance"},
                new(){ Label="交通", Prop="Transportation"},
                new(){ Label="餐补", Prop="Meal"},
            };

            List<Salary> data = new List<Salary>();

            data.Add(new Salary() { Name = "张三", BaseSalary = 8000, FullAttendance = 200, Meal = 300, Transportation = 300 });
            data.Add(new Salary() { Name = "李四", BaseSalary = 9000, FullAttendance = 300, Meal = 300, Transportation = 300 });
            data.Add(new Salary() { Name = "王五", BaseSalary = 1000, FullAttendance = 400, Meal = 300, Transportation = 300 });

            EasyExport.CreateSheet(wb, columns, data);

            using (MemoryStream ms = new())
            {
                wb.Write(ms);
                wb.Close();
                return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}.xlsx");
            }
        }

        /// <summary>
        /// 格式化列（写法一），例如：多列相加，转换时间格式，求和等
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IActionResult Demo4()
        {
            IWorkbook wb = new XSSFWorkbook();


            //利用 Format 格式化列
            List<ExportColumn> columns = new()
            {
                new(){ Label="姓名", Prop="Name"},
                new(){ Label="底薪", Prop="BaseSalary"},
                new(){ Label="全勤", Prop="FullAttendance"},
                new(){ Label="交通", Prop="Transportation"},
                new(){ Label="餐补", Prop="Meal"},
                new(){
                    Label="合计",
                    //第一个参数是当前行数据，第二个参数是当前行索引
                    Format = (row,rowIndex)=>{
                        Salary salary = (Salary)row;
                        return (double)salary.BaseSalary+(double)salary.FullAttendance+(double)salary.Transportation+(double)salary.Meal;
                    }
                },
            };

            List<Salary> data = new List<Salary>();

            data.Add(new Salary() { Name = "张三", BaseSalary = 8000, FullAttendance = 200, Meal = 300, Transportation = 300 });
            data.Add(new Salary() { Name = "李四", BaseSalary = 9000, FullAttendance = 300, Meal = 300, Transportation = 300 });
            data.Add(new Salary() { Name = "王五", BaseSalary = 1000, FullAttendance = 400, Meal = 300, Transportation = 300 });

            EasyExport.CreateSheet(wb, columns, data);

            //如果format 比较复杂，并且个数较多，或想使用泛型， 可以单独设置format字典， 这时表头中的 Prop 字段必填，key为 Prop，效果同上

            //Dictionary<string, Func<Salary, int, string>> formatDic = new Dictionary<string, Func<Salary, int, string>>();
            //formatDic.Add("Total", (row, rowIndex) =>
            //{
            //    return (row.BaseSalary + row.FullAttendance + row.Transportation + row.Meal).ToString();
            //});
            //MuziExport.CreateSheet(wb, columns, data, formatDic);


            using (MemoryStream ms = new())
            {
                wb.Write(ms);
                wb.Close();
                return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}.xlsx");
            }
        }

        /// <summary>
        /// 格式化列（写法二），使用泛型
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IActionResult Demo5()
        {
            IWorkbook wb = new XSSFWorkbook();

            List<ExportColumn> columns = new()
            {
                new(){ Label="姓名", Prop="Name"},
                new(){ Label="底薪", Prop="BaseSalary"},
                new(){ Label="全勤", Prop="FullAttendance"},
                new(){ Label="交通", Prop="Transportation"},
                new(){ Label="餐补", Prop="Meal"},
                new(){ Label="合计", Prop="Total"},
            };

            List<Salary> data = new List<Salary>();

            data.Add(new Salary() { Name = "张三", BaseSalary = 8000, FullAttendance = 200, Meal = 300, Transportation = 300 });
            data.Add(new Salary() { Name = "李四", BaseSalary = 9000, FullAttendance = 300, Meal = 300, Transportation = 300 });
            data.Add(new Salary() { Name = "王五", BaseSalary = 1000, FullAttendance = 400, Meal = 300, Transportation = 300 });

            //如果format 比较复杂，并且个数较多，或想使用泛型， 可以单独设置format字典， 这时表头中的 Prop 字段必填，key为 Prop
            //表头内format 优先级高于 字典format

            Dictionary<string, Func<Salary, int, object>> formatDic = new Dictionary<string, Func<Salary, int, object>>();
            formatDic.Add("Total", (row, rowIndex) =>
            {
                return ((double)row.BaseSalary + (double)row.FullAttendance + (double)row.Transportation + row.Meal).ToString();
            });

            EasyExport.CreateSheet(wb, columns, data, formatDic);


            using (MemoryStream ms = new())
            {
                wb.Write(ms);
                wb.Close();
                return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}.xlsx");
            }
        }

        /// <summary>
        /// 大量数据 10 万
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IActionResult Demo6()
        {
            // Muzi.ExcelExport 只是针对 IWorkBook 和  ISheet 的处理

            //不同的需求请选择不同的 IWorkBook实现

            //NPOI 提供了以下三种

            //**HSSFWorkbook、XSSFWorkbook 和 SXSSFWorkbook**

            //简而言之，

            //HSSFWorkbook 针对Excel2003 ，有65535的条数限制

            //XSSFWorkbook 针对Excel2007,     没有条数限制，但是内存占用可能大

            //SXSSFWorkbook 是XSSFWorkbook 的内存改进版，但是有操作局限性

            //具体情况请自行查阅

            //下面以10万条数据导出为例

            IWorkbook wb = new SXSSFWorkbook();

            List<ExportColumn> columns = new()
            {
                new(){ Label="姓名", Prop="Name",Width=200},
                new(){ Label="底薪", Prop="BaseSalary",Width=200},
                new(){ Label="全勤", Prop="FullAttendance",Width=200},
                new(){ Label="交通", Prop="Transportation",Width=200},
                new(){ Label="餐补", Prop="Meal",Width=200},
                new(){ Label="时间", Prop="CreateTime",Width=300},
            };

            List<Salary> data = new List<Salary>();

            for (int i = 0; i < 10000 * 10; i++)
            {
                data.Add(new Salary() { Name = "张三", BaseSalary = 8000, FullAttendance = 200, Meal = 300, Transportation = 300, CreateTime = DateTime.Now });
            }

            EasyExport.CreateSheet(wb, columns, data);

            using (MemoryStream ms = new())
            {
                wb.Write(ms);
                wb.Close();
                return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}.xlsx");
            }
        }

        /// <summary>
        /// 全局设置 和 自定义钩子
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IActionResult Demo7()
        {
            IWorkbook wb = new XSSFWorkbook();

            List<ExportColumn> columns = new()
            {
                new(){ Label="姓名", Prop="Name"},
                new(){ Label="底薪", Prop="BaseSalary"},
                new(){ Label="全勤", Prop="FullAttendance"},
                new(){ Label="交通", Prop="Transportation"},
                new(){ Label="餐补", Prop="Meal"},
            };

            List<Salary> data = new List<Salary>();

            data.Add(new Salary() { Name = "张三", BaseSalary = 8000, FullAttendance = 0, Meal = 300, Transportation = 300 });
            data.Add(new Salary() { Name = "李四", BaseSalary = 8000, FullAttendance = 0, Meal = 300, Transportation = 300 });
            data.Add(new Salary() { Name = "王五", BaseSalary = 8000, FullAttendance = 0, Meal = 300, Transportation = 300 });
            data.Add(new Salary() { Name = "赵六", BaseSalary = 8000, FullAttendance = 0, Meal = 300, Transportation = 300 });
            data.Add(new Salary() { Name = "孙七", BaseSalary = 8000, FullAttendance = 0, Meal = 300, Transportation = 300 });

            ICellStyle headerStyle = EasyExport.GetDefaultHeaderStyle(wb);

            ExportOptions exportOptions = new ExportOptions();
            //表头行高
            exportOptions.HeaderHeight = 400;
            //数据行高
            exportOptions.DataRowHeight = 300;
            //表头样式
            exportOptions.HeaderStyle = headerStyle;
            //数据样式
            exportOptions.DataStyle = null;

            //另外提供两个自定义钩子，分别是  RowCreateAfter 和  CellCreateAfter

            //给索引为1的行高设置为500；
            exportOptions.RowCreateAfter = (row, rowIndex) => { if (rowIndex == 1) row.Height = 800; };

            //给行列相等的单元格背景色
            exportOptions.CellCreateAfter = (cell, colIndex, row, rowIndex) =>
            {
                if (colIndex == rowIndex)
                {
                    ICellStyle cellStyle = wb.CreateCellStyle();

                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;

                    cell.CellStyle = cellStyle;
                }

            };

            EasyExport.CreateSheet(wb, columns, data, exportOptions: exportOptions);

            using (MemoryStream ms = new())
            {
                wb.Write(ms);
                wb.Close();
                return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}.xlsx");
            }
        }
    }

}
