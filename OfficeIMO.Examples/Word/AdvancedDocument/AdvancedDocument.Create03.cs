using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Color = SixLabors.ImageSharp.Color;
using Path = System.IO.Path;

namespace OfficeIMO.Examples.Word;
internal static partial class AdvancedDocument {

    public static void Example_AdvancedWord3(string templatePath, string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating advanced document");
        string tempPath = Path.Combine(templatePath, "AdvancedDocument3.docx");
        string filePath = Path.Combine(folderPath, "AdvancedDocument3.docx");

        using (WordDocument document = WordDocument.Load(tempPath)) {

            ReplaceMark(document);
            AddTable(document);
            AddImage(document);
            AddChapter(document);
            AddingCharts(document);
            document.Save(filePath, true);
        }
    }


    private static void ReplaceMark(WordDocument document) {
        var dict = new Dictionary<string, string>() { };
        dict.Add("ProjectName", "测试项目ABC");
        dict.Add("ProjectId", "NO2023-0704-01");
        dict.Add("AppVer", "BPA-2023.4.0");
        dict.Add("InitBlock", "1000,000");
        dict.Add("Relaxation.Press", "1000");

        foreach (var item in dict) {
            document.FindAndReplace($"{{{item.Key}}}", item.Value);
        }

    }

    public static void AddTable(WordDocument document) {
        var p = document.AddParagraph().SetStyleId("1").SetText("添加表格");
        p = document.AddParagraph().SetStyleId("2").SetText("简单表格");

        var table = document.AddTable(5, 2);
        table.Alignment = TableRowAlignmentValues.Center;
        table.Width = 3000;
        table.WidthType = TableWidthUnitValues.Pct;
        table.Alignment= TableRowAlignmentValues.Center;
        table.ColumnWidth = new List<int>() { 1500, 3500 };

        for (int i = 0; i < 5; i++) {
            table.Rows[i].FirstCell.ShadingFillColor = Color.FromRgb(117, 117, 117);
            var cell = table.Rows[i].FirstCell.Paragraphs[0];
            cell.Text = $"项目名称{i + 1}";
            cell.FontSize = 15;
            cell.FontFamily = "SimSun";
            cell.ParagraphAlignment = JustificationValues.Center;
            cell = table.Rows[i].Cells[1].Paragraphs[0];
            cell.SetFontFamily("SimSun");
            cell.SetFontSize(15);
            cell.ParagraphAlignment = JustificationValues.Center;
            cell.Text = $"值{i + 1}";
            if (i % 2 == 0) {
                cell.Color = Color.FromRgb(255, 0, 0);
            }
        }

        p = document.AddParagraph().SetStyleId("2").SetText("合并表格");

        table = document.AddTable(7, 9);
        table.Width = 5000;
        table.WidthType = TableWidthUnitValues.Auto;
       
        for (int i = 0; i < 7; i++) {
            var row = table.Rows[i];
            if (i == 0) {
                row.Cells[0].Paragraphs[0].Text = "所属建筑";
                row.Cells[1].Paragraphs[0].Text = "所属楼层";
                row.Cells[2].Paragraphs[0].Text = "设备名称";
                row.Cells[3].Paragraphs[0].Text = "进风温度(℃)";
                row.Cells[4].Paragraphs[0].Text = "排风温度(℃)";
                row.Cells[5].Paragraphs[0].Text = "环境温度(℃)";
                row.Cells[6].Paragraphs[0].Text = "进风温度限值(℃)";
                row.Cells[7].Paragraphs[0].Text = "是否超过限值";
                row.Cells[8].Paragraphs[0].Text = "返混率(%)";
                for (int j = 0; j < 9; j++) {
                    row.Cells[j].ShadingFillColor = Color.FromRgb(117, 117, 117);
                }
            } else {
                row.Cells[0].Paragraphs[0].Text = "bldg-1";
                row.Cells[1].Paragraphs[0].Text = "建模_标高 8";
                row.Cells[2].Paragraphs[0].Text = $"室外机11-{i}";
                row.Cells[3].Paragraphs[0].Text = "29.92";
                row.Cells[4].Paragraphs[0].Text = "48.22";
                row.Cells[5].Paragraphs[0].Text = "29.9";
                row.Cells[6].Paragraphs[0].Text = "30";
                if (i%2==0) {
                    row.Cells[7].Paragraphs[0].Text = "是";
                } else {
                    row.Cells[7].Paragraphs[0].Text = "否";
                }
                row.Cells[8].Paragraphs[0].Text = "0.1";
            }

            for (int j = 0; j < 9; j++) {
                row.Cells[j].Paragraphs[0].ParagraphAlignment = JustificationValues.Center;
                row.Cells[j].VerticalAlignment = TableVerticalAlignmentValues.Center;
                if (i%2==0&&i>0) {
                    row.Cells[j].Paragraphs[0].Color = Color.FromRgb(255, 0, 0);
                }
            }
            table.Rows[1].Cells[0].MergeVertically(7, false);

        }



    }

    private static void AddImage(WordDocument document) {
        var p = document.AddParagraph().SetStyleId("1").SetText("添加图片");

        var imageFolder = Path.Combine(Directory.GetCurrentDirectory(), "Images");
        var image = Path.Combine(imageFolder, "Kulek.jpg");

        document.AddParagraph("添加JPG图片").SetAlignment(JustificationValues.Center);
        var paragraph3 = document.AddParagraph();
       
        paragraph3.AddImage(image, 500, 500);

        document.AddParagraph("添加PNG图片").SetAlignment(JustificationValues.Center);
        var paragraph4 = document.AddParagraph().SetAlignment(JustificationValues.Center);
        image= Path.Combine(imageFolder, "EvotecLogo.png");
        paragraph4.AddImage(image, 100, 100);

    }

    private static void AddChapter(WordDocument document) {

    }

    private static void AddingCharts(WordDocument document) {

        document.AddParagraph().SetStyleId("1").SetText("添加图表");

        document.AddParagraph().SetStyleId("2").SetText("柱状图");
        var random = Random.Shared.Next(10);
        List<string> categories = Enumerable.Range(1,12).Select(i => $"{i}月").ToList();
        List<int> values1=Enumerable.Range(1,12).Select(i=>(i%10) * Random.Shared.Next(10)).ToList();
        List<int> values2 = Enumerable.Range(1, 12).Select(i => (i % 10) * Random.Shared.Next(10)).ToList();
        List<int> values3 = Enumerable.Range(1, 12).Select(i => (i % 10) * Random.Shared.Next(10)).ToList();
        List<int> values4 = Enumerable.Range(1, 12).Select(i => (i % 10) * Random.Shared.Next(10)).ToList();
       
        var barChart1 = document.AddBarChart();
        barChart1.AddCategories(categories);
        barChart1.AddChartBar("冷负荷(Kw.h)", values1, Color.Brown);
        barChart1.AddChartBar("热负荷(kW.h)", values2, Color.Green);
        barChart1.AddChartBar("生活热水负荷(kW.h)", values3, Color.DarkGoldenrod);
        barChart1.AddChartBar("新风热回收冷量(kW.h)", values4, Color.GreenYellow);
        barChart1.AddLegend(LegendPositionValues.Top);
        barChart1.BarGrouping = BarGroupingValues.Clustered;
        barChart1.BarDirection = BarDirectionValues.Column;

        document.AddParagraph().SetStyleId("2").SetText("面积图");

        var areaChart = document.AddAreaChart("面积图");
        areaChart.AddCategories(categories);
        areaChart.AddChartArea("冷负荷(Kw.h)", values1, Color.Brown);
        areaChart.AddChartArea("热负荷(kW.h)", values2, Color.Green);
        areaChart.AddChartArea("生活热水负荷(kW.h)", values3, Color.DarkGoldenrod);
        areaChart.AddChartArea("新风热回收冷量(kW.h)", values4, Color.GreenYellow);
        areaChart.AddLegend(LegendPositionValues.Top);
      


        document.AddParagraph().SetStyleId("2").SetText("折线图");

        var lineChart2 = document.AddLineChart();
        lineChart2.AddChartAxisX(categories);
        lineChart2.AddLegend(LegendPositionValues.Bottom);
        lineChart2.AddChartLine("冷负荷(Kw.h)", values1, Color.Brown);
        lineChart2.AddChartLine("热负荷(kW.h)", values2, Color.Green);
        lineChart2.AddChartLine("生活热水负荷(kW.h)", values3, Color.DarkGoldenrod);
        lineChart2.AddChartLine("新风热回收冷量(kW.h)", values4, Color.GreenYellow);


        document.AddParagraph().SetStyleId("2").SetText("饼图");

        var pieChart = document.AddPieChart();
        pieChart.AddCategories(categories);
        pieChart.AddLegend(LegendPositionValues.TopRight);
        pieChart.AddChartPie("全年负荷比率", values1);



    }


}


