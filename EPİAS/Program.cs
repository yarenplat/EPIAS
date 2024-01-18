using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        // JSON verisini bir dosyadan oku veya string olarak elde et
        string jsonVeri = File.ReadAllText("C:\\Users\\YarenPOLAT\\Desktop\\response.json");

        // JSON verisini C# nesnesine dönüştür
        var veriNesnesi = JsonConvert.DeserializeObject<VeriNesnesi>(jsonVeri);

        // Excel dosyası oluştur
        using (var package = new ExcelPackage())
        {
            // Excel sayfası oluştur
            var worksheet = package.Workbook.Worksheets.Add("VeriSayfasi");

            // Başlık satırını ekleyerek başla
            YazdirBasliklar(worksheet, 1, 1);

            // JSON verisini Excel'e yaz
            YazdirVeri(veriNesnesi.items, worksheet, 2, 1);

            // Excel dosyasını kaydet
            package.SaveAs(new FileInfo("sonuc.xlsx"));
        }

        Console.WriteLine("İşlem tamamlandı. Excel dosyası oluşturuldu.");
    }

    static void YazdirBasliklar(ExcelWorksheet worksheet, int satir, int sutun)
    {
        // Excel sayfasına başlık satırını yaz
        var basliklar = new List<string>
        {
            "Date", "Hour", "Total", "NaturalGas", "DammedHydro", "Lignite",
            "River", "ImportCoal", "Wind", "Sun", "FuelOil", "Geothermal",
            "AsphaltiteCoal", "BlackCoal", "Biomass", "Naphta", "Lng",
            "ImportExport", "WasteHeat"
        };

        for (int i = 0; i < basliklar.Count; i++)
        {
            worksheet.Cells[satir, sutun + i].Value = basliklar[i];
        }
    }

    static void YazdirVeri(List<Item> veri, ExcelWorksheet worksheet, int satir, int sutun)
    {
        // Her bir JSON öğesini Excel satırına yaz
        foreach (var item in veri)
        {
            worksheet.Cells[satir, sutun].Value = item.date;
            worksheet.Cells[satir, sutun + 1].Value = item.hour;
            worksheet.Cells[satir, sutun + 2].Value = item.total;
            worksheet.Cells[satir, sutun + 3].Value = item.naturalGas;
            worksheet.Cells[satir, sutun + 4].Value = item.dammedHydro;
            worksheet.Cells[satir, sutun + 5].Value = item.lignite;
            worksheet.Cells[satir, sutun + 6].Value = item.river;
            worksheet.Cells[satir, sutun + 7].Value = item.importCoal;
            worksheet.Cells[satir, sutun + 8].Value = item.wind;
            worksheet.Cells[satir, sutun + 9].Value = item.sun;
            worksheet.Cells[satir, sutun + 10].Value = item.fueloil;
            worksheet.Cells[satir, sutun + 11].Value = item.geothermal;
            worksheet.Cells[satir, sutun + 12].Value = item.asphaltiteCoal;
            worksheet.Cells[satir, sutun + 13].Value = item.blackCoal;
            worksheet.Cells[satir, sutun + 14].Value = item.biomass;
            worksheet.Cells[satir, sutun + 15].Value = item.naphta;
            worksheet.Cells[satir, sutun + 16].Value = item.lng;
            worksheet.Cells[satir, sutun + 17].Value = item.importExport;
            worksheet.Cells[satir, sutun + 18].Value = item.wasteheat;

            satir++;
        }
    }
}

public class Item
{
    public string date { get; set; }
    public string hour { get; set; }
    public double total { get; set; }
    public double naturalGas { get; set; }
    public double dammedHydro { get; set; }
    public double lignite { get; set; }
    public double river { get; set; }
    public double importCoal { get; set; }
    public double wind { get; set; }
    public double sun { get; set; }
    public double fueloil { get; set; }
    public double geothermal { get; set; }
    public double asphaltiteCoal { get; set; }
    public double blackCoal { get; set; }
    public double biomass { get; set; }
    public double naphta { get; set; }
    public double lng { get; set; }
    public double importExport { get; set; }
    public double wasteheat { get; set; }
}

public class Page
{
    public int number { get; set; }
    public int size { get; set; }
    public int total { get; set; }
    public Sort sort { get; set; }
}

public class Sort
{
    public string field { get; set; }
    public string direction { get; set; }
}

public class Totals
{
    public double fueloilTotal { get; set; }
    public object gasOilTotal { get; set; }
    public double blackCoalTotal { get; set; }
    public double ligniteTotal { get; set; }
    public double geothermalTotal { get; set; }
    public double naturalGasTotal { get; set; }
    public double riverTotal { get; set; }
    public double dammedHydroTotal { get; set; }
    public double lngTotal { get; set; }
    public double biomassTotal { get; set; }
    public double naphtaTotal { get; set; }
    public double importCoalTotal { get; set; }
    public double asphaltiteCoalTotal { get; set; }
    public double windTotal { get; set; }
    public object nuclearTotal { get; set; }
    public double sunTotal { get; set; }
    public double importExportTotal { get; set; }
    public double wasteheatTotal { get; set; }
    public double totalTotal { get; set; }
}

public class VeriNesnesi
{
    public List<Item> items { get; set; }
    public Page page { get; set; }
    public Totals totals { get; set; }
}
