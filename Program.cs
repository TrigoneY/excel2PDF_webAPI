using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.Extensions.FileProviders;
using System;
using System.Diagnostics;
using System.Reflection.PortableExecutable;
using System.Text.Json;
using System.Xml;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
//builder.Services.AddEndpointsApiExplorer();
//builder.Services.AddSwaggerGen();
builder.Services.AddControllers();
object cmdlock = new object();

var app = builder.Build();

// Configure the HTTP request pipeline.
//if (app.Environment.IsDevelopment())
//{
//    app.UseSwagger();
//    app.UseSwaggerUI();
//}

app.UseHttpsRedirection();

app.MapGet("/excel2pdf/list", (FileType type) =>
{
    string tail = "";
    switch (type)
    {
        case FileType.PDF:
            tail = "pdf";
            break;
        case FileType.XLSX:
            tail = "xlsx";
            break;
    }
    string[] files = Directory.GetFiles("./data", "*." + tail).Select(filePath => Path.GetFileName(filePath)).ToArray();
    string jsonResult = JsonSerializer.Serialize(files);
    return jsonResult;
});
//.WithName("list-exceluploaded")
//.WithOpenApi();

app.MapGet("/excel2pdf/update", (string key, string value) =>
{
    ProcessStartInfo startInfo = new ProcessStartInfo();
    startInfo.CreateNoWindow = true;
    startInfo.UseShellExecute = false;
    startInfo.FileName = "CMD.exe";
    startInfo.WorkingDirectory = @"./data";
    string cmd_arg0 = "/C D:\\soft\\LibreOffice\\program\\soffice --headless toPDFtemp.ods "; // remain a space at tail
    string cmd_arg1 = $"\"macro://./Standard.calc2PDF.updateByKeValue(\"{key}\", \"{value}\")\" ";
    startInfo.Arguments = cmd_arg0 + cmd_arg1 ;
    bool cmdResult;
    //lock (cmdlock)
    //{
        using (Process process = new Process())
        {
            process.StartInfo = startInfo;
            process.Start();
            cmdResult = process.WaitForExit(9999);
            process.Close();
        }
    //}
    return cmdResult;
});

app.MapGet("/excel2pdf", (string xlsxfilename) =>
{
    var result = new List<string>();
    ProcessStartInfo startInfo = new ProcessStartInfo();
    startInfo.CreateNoWindow = true;
    startInfo.UseShellExecute = false;
    //startInfo.RedirectStandardOutput = true; //unwork 
    //startInfo.RedirectStandardError = true; //unwork 
    startInfo.FileName = "CMD.exe";
    startInfo.WorkingDirectory = @"./data";
    //startInfo.EnvironmentVariables["PATH"] = "D:\\soft\\LibreOffice\\program;"; //unwork 
    //startInfo.WindowStyle = ProcessWindowStyle.Hidden;
    string cmd_arg0 = "/C D:\\soft\\LibreOffice\\program\\soffice --headless toPDFtemp.ods "; // remain a space at tail
    //string cmd_arg11 = $"macro://./Standard.calc2PDF.updateFilename(\"{xlsxfilename}\") ";
    string cmd_arg1 = $"\"macro://./Standard.calc2PDF.updateByKeValue(\"FileName\", \"{xlsxfilename}\")\" ";
    string cmd_arg2 = "\"macro://./Standard.calc2PDF.Main\" ";
    startInfo.Arguments = cmd_arg0 + cmd_arg1 + cmd_arg2;
    //startInfo.Arguments = "/C D:\\soft\\LibreOffice\\program\\soffice --headless toPDFtemp.ods \"vnd.sun.star.script:Standard.calc2PDF.test?language=Basic&location=document\"";

    lock (cmdlock)
    {
        using (Process process = new Process())
        {
            process.StartInfo = startInfo;
            //process.OutputDataReceived += (sender, data) => result.Add(data.Data); //unwork 
            //process.ErrorDataReceived += (sender, data) => result.Add(data.Data); //unwork 
            process.Start();
            process.WaitForExit(9999);
            //process.BeginOutputReadLine(); //unwork 
            //process.BeginErrorReadLine(); //unwork 
            process.Close();
        }
    }
    
    return result;
});
//.WithName("convert-excel2PDF")
//.WithOpenApi();

//https://learn.microsoft.com/en-us/aspnet/core/fundamentals/static-files?view=aspnetcore-8.0
app.UseStaticFiles(new StaticFileOptions
{
    FileProvider = new PhysicalFileProvider(
        Path.Combine(Directory.GetCurrentDirectory(), "data")),
    RequestPath = "/StaticFile"
});

app.UseRouting();
app.MapControllers();

app.Run();

public enum FileType
{
    PDF = 1,
    XLSX = 2,
    XLS = 3,
}




