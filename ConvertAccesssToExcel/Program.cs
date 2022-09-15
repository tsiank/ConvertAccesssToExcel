using System.Data;
using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using System.Data.OleDb;


Console.WriteLine("1、请输入mdb文件所在目录：");
var mdbPath = Console.ReadLine();
Console.WriteLine("2、请输入保存目录：");
var savePath = Console.ReadLine();
Console.WriteLine("转换中...");

var files = Directory.GetFiles(mdbPath);
foreach (var file in files)
{
    ConvertAccessTableToExcel(file, savePath);
}
Console.WriteLine("完成！");
Console.ReadKey();

//转换单个access文件所有表 
void ConvertAccessTableToExcel(string mdbFile, string savePath)
{

    var config = new OpenXmlConfiguration()
    {
        TableStyles = MiniExcelLibs.OpenXml.TableStyles.None,
        EnableConvertByteArray = true
    };

    string strConn = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={mdbFile}";
    using var conn = new OleDbConnection(strConn);
    conn.Open();

    var schema = conn.GetSchema("Tables");
    var userTables = schema.AsEnumerable().Where(s => s.Field<string>("TABLE_TYPE") == "TABLE").Select(s => s.Field<string>("TABLE_NAME"));

    var baseName = Path.GetFileNameWithoutExtension(mdbFile);

    foreach (var tableName in userTables)
    {
        var cmd = new OleDbCommand("select * from " + tableName, conn);

        //var adapter = new OleDbDataAdapter(cmd);
        //var dt = new DataTable();
        //adapter.Fill(dt);

        var reader = cmd.ExecuteReader();
        var path = @$"{savePath}\{baseName}_{tableName}.xlsx";
        //MiniExcel.SaveAs(path, reader, configuration: config, overwriteFile: true);
        MiniExcel.SaveAs(path, reader, overwriteFile: true);
    }
}
