# 前言

在使用NPOI  导出Excel 的过程中，对于表头合并，列的映射，以及自定义处理 并不是很方便

突发奇想，能不能让NPOI的导出  像前端画表格一样方便

类似于 BootStrapTable   ,只需要定义好表头，将数据进行填充，自动完成Excel处理

于是乎 Excel.EasyExport诞生了!

# 重要

> **以下示例均采用 C# 9.0 新语法糖  new()   的简略写法，c# 9之前请自行补全**



# 项目地址

github:https://github.com/TomatoDAIXIA/Excel.EasyExport

gitee:https://gitee.com/tomato23132313/Excel.EasyExport



# 初体验

导出一个带多级合并表头的Excel

效果如下：

![在这里插入图片描述](https://img-blog.csdnimg.cn/7df0600770d34c458413ac42418b5321.png#pic_center)

你只需要编写：

```c#
IWorkbook wb = new XSSFWorkbook();
//定义表头
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
//获取数据
List<Salary> data = new() {
    new() { Name = "张三",
        BaseSalary = 8000,
        FullAttendance = 200,
        Meal = 300,
        Transportation = 300,
        CreateTime = DateTime.Now
    }
};
//创建Sheet
var sheet = EasyExport.CreateSheet(wb, columns, data);

```

# 详细文档

## 安装



NuGet 包搜索  Excel.EasyExport

![在这里插入图片描述](https://img-blog.csdnimg.cn/35412f2294ad4affb49ae90df010c4de.png#pic_center)


或者Nuget 控制台

```
Install-Package Excel.EasyExport
```



## 用法

以下示例实体均来自于：

```c#
public class Salary
{
    public string Name { get; set; }
    public decimal BaseSalary { get; set; }
    public decimal FullAttendance { get; set; }
    public decimal Transportation { get; set; }
    public decimal Meal { get; set; }

    public DateTime? CreateTime { get; set; }

    public bool IsFullAttendance { get; set; }
}

```



### 自动映射字段和类型



> 属性名和实体字段对应，区分大小写
>
> Prop除了直接写字符串，还可使用 nameof() 语法，例如 Prop=nameof(Salary.CreateTime)
>
> 除映射字段外，还提供4种类型自动映射
>
> | 转换类型 | 原类型                                                       |
> | -------- | ------------------------------------------------------------ |
> | double   | decimal、byte、sbyte、short、ushort 、 int、uint、long、ulong、float、double |
> | bool     | bool                                                         |
> | DateTime | DateTime                                                     |
> | string   | 其他                                                         |
>
> 



```c#
 IWorkbook wb = new XSSFWorkbook();
 //Prop 对应实体的字段名，区分大小写
 List<ExportColumn> columns = new()
 {
     new(){ Label="姓名", Prop="Name"},
     new(){ Label="底薪", Prop="BaseSalary"},
     new(){ Label="全勤", Prop="FullAttendance",},
     new(){ Label="交通", Prop="Transportation"},
     new(){ Label="餐补", Prop="Meal",},
     new(){ Label="创建时间", Prop=nameof(Salary.CreateTime)},
 };

 List<Salary> data = new List<Salary>();
 data.Add(new Salary()
 {
     Name = "张三",
     BaseSalary = 8000,
     FullAttendance = 200,
     Meal = 300,
     Transportation = 300,
     CreateTime = DateTime.Now
 });

 var sheet = EasyExport.CreateSheet(wb, columns, data);

```



### 合并表头

> 正如文档开头，这里借鉴了前端的使用方式，通过父子级关系自动嵌套合并
>
> 值得注意的是，只有叶子节点的Prop才会生效

```c#
IWorkbook wb = new XSSFWorkbook();
//定义表头
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
//获取数据
List<Salary> data = new() {
    new() { Name = "张三",
        BaseSalary = 8000,
        FullAttendance = 200,
        Meal = 300,
        Transportation = 300,
        CreateTime = DateTime.Now
    }
};
//创建Sheet
var sheet = EasyExport.CreateSheet(wb, columns, data);
```



### 指定列样式

> 通常，我们需要给指定列定制样式，
>
> 例如：  姓名一列 加粗靠左对齐
>
> 你只需要像编写 bootstrapTable 一样简单的使用他
>
> Style 提供了一个 ICellStyle 返回类型的委托，这完全是NPOI内置的样式，你可以随意定制你的样式
>
> 每一列的 Style 只会创建一次，并不会重复创建



```c#
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

```

效果：

![在这里插入图片描述](https://img-blog.csdnimg.cn/a230c9d329644b36abf9047a90d56ac7.png#pic_center)




### 格式化列

> 实际上，有时我们需要的不是原始字段的值，甚至需要多加一列虚拟列
>
> 类如：增加一列合计、给底薪乘上一个系数、将 True\false 转换为是\否 等等
>
> 这里以增加一列合计为例：



> Format 提供了 两个可使用的参数，并返回一个 object
>
> 第一个入参：当前行数据 object 类型
>
> 第二个入参：当前行索引  int 类型
>
> 返回值：object 类型



```c#
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

EasyExport.CreateSheet(wb, columns, data);

```

> 如果你的格式列较多，或者你想使用泛型，又或者你想保持代码的清爽
>
> 这里提供第二种格式化方式  外部字典Fromat
>
> 值得注意的是，这种方式Prop 必填

```c#
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

//如果format 比较复杂，并且个数较多，或想使用泛型， 可以单独设置format字典， 这时表头中的 Prop 字段必填，key为 Prop
//表头内format 优先级高于 字典format
Dictionary<string, Func<Salary, int, object>> formatDic = new Dictionary<string, Func<Salary, int, object>>();
formatDic.Add("Total", (row, rowIndex) =>
{
    return ((double)row.BaseSalary + (double)row.FullAttendance + (double)row.Transportation + row.Meal).ToString();
});

EasyExport.CreateSheet(wb, columns, data, formatDic);

```

### 列字段说明

> 创建表头需要用到  ExportColumn 类



| 列名     | 类型                      | 描述                          |
| -------- | ------------------------- | ----------------------------- |
| Label    | string                    | 表题文本                      |
| Prop     | string                    | 字段名                        |
| Width    | int                       | 列宽，详见 **关于列宽的说明** |
| Style    | Func<ICellStyle>          | 列样式                        |
| Format   | Func<object, int, object> | 格式化列                      |
| Children | List<ExportColumn>        | 子集                          |

### 关于列宽的说明

如果列中的Width 不赋值，默认为0，这时将开启自动列宽，Width 只在叶子节点中有效

自动列宽在 NPOI AutoSizeColumn 方法的基础上自动增加了300 的列宽，这使的列宽的效果更好

但值得注意的是，在大量数据下，AutoSizeColumn  将会严重影响速度和性能

特做此处理： 当数据量  > 5000 行时，自动禁用  AutoSizeColumn

你也可以在全局设置中，使用exportOptions 参数，一开始就禁用 AutoSizeColumn，详见 **全局设置**



### 全局设置

> 除了上面提到的功能，Muzi.ExcelExport 还提供了几种实用的全局设置
>
> ExportOptions 类



| 字段            | 类型                          | 默认值 | 描述                                                         |
| --------------- | ----------------------------- | ------ | ------------------------------------------------------------ |
| HeaderHeight    | short                         | 400    | 表头行高                                                     |
| DataRowHeight   | short                         | 300    | 数据行高                                                     |
| HeaderStyle     | ICellStyle                    | null   | 表头单元格样式                                               |
| DataStyle       | ICellStyle                    | null   | 数据单元格样式                                               |
| RowCreateAfter  | Action<IRow, int>             | null   | IRow 创建后的钩子函数,可实现自定义逻辑，第一个参数：IRow,第二个参数 行索引 |
| CellCreateAfter | Action<ICell, int, IRow, int> | null   | ICell 创建后的钩子函数，参数依次为 ICell对象，列索引， IRow对象，行索引， |
| AutoColSize     | bool                          | true   | 是否开启自动列宽，特别提醒：当数据条数 >5000 时失效，需要自行设置列宽 |



```c#
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

```



效果


![在这里插入图片描述](https://img-blog.csdnimg.cn/8bc457a524ea4a45a7eefeb95d7a16ef.png#pic_center)




### 关于大量数据的说明



> Muzi.ExcelExport 只是针对 IWorkBook 和  ISheet 的处理
>
> 不同的需求请选择不同的 IWorkBook实现
>
> NPOI 提供了以下三种
>
> **HSSFWorkbook、XSSFWorkbook 和 SXSSFWorkbook**
>
> 简而言之，
>
> HSSFWorkbook 针对Excel2003 ，有65535的条数限制
>
> XSSFWorkbook 针对Excel2007,     没有条数限制，但是内存占用可能大
>
> SXSSFWorkbook 是XSSFWorkbook 的内存改进版，但是有操作局限性
>
> 具体情况请自行查阅
>
> 下面以10万条数据导出为例



```c#
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
    data.Add(new Salary(){ 
        Name = "张三", 
        BaseSalary = 8000, 
        FullAttendance = 200,
        Meal = 300, 
        Transportation = 300, 
        CreateTime = DateTime.Now 
    });
}

EasyExport.CreateSheet(wb, columns, data);
```



### 保存

> 由于保存业务不尽相同，这里不做封装，只提2种写法仅供参考



保存到文件

```C#
using (var fs = File.Open(path, FileMode.Create, FileAccess.Write))
    workbook.Write(fs);
```



 farmework   .net core  webApi 导出

```c# 
using (MemoryStream ms = new())
{
    workbook.Write(ms);
    workbook.Close();
    return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "test.xlsx");
}
```
