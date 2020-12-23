using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace InventarioToWooCommerce
{
    class Program
    {
        static void Main(string[] args)
        {
            //Make Changes for other page
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string fileDirectory = $"{Path.GetFullPath(@"..\..\..\..\")}Inventario_MBD.xlsx";


            using (var package = new ExcelPackage(new FileInfo(fileDirectory)))
            {
                //Adding new sheet
                package.Workbook.Worksheets.Add("UploadData");
                var rawSheet = package.Workbook.Worksheets["Ellas"];
                var newSheet = package.Workbook.Worksheets["UploadData"];

                //Adding Colums Headers
                newSheet.Cells[$"A1"].Value = "type";
                newSheet.Cells[$"B1"].Value = "SKU";
                newSheet.Cells[$"C1"].Value = "Name";
                newSheet.Cells[$"D1"].Value = "Published";
                newSheet.Cells[$"E1"].Value = "Is featured?";
                newSheet.Cells[$"F1"].Value = "Visibility in catalog";
                newSheet.Cells[$"G1"].Value = "Short description";
                newSheet.Cells[$"H1"].Value = "Description";
                newSheet.Cells[$"I1"].Value = "In stock?";
                newSheet.Cells[$"J1"].Value = "Stock";
                newSheet.Cells[$"K1"].Value = "Position";
                newSheet.Cells[$"L1"].Value = "Attribute 1 name";
                newSheet.Cells[$"M1"].Value = "Attribute 1 value(s)";
                newSheet.Cells[$"N1"].Value = "Attribute 1 visible";
                newSheet.Cells[$"O1"].Value = "Attribute 1 global";
                newSheet.Cells[$"P1"].Value = "Attribute 2 name";
                newSheet.Cells[$"Q1"].Value = "Attribute 2 value(s)";
                newSheet.Cells[$"R1"].Value = "Attribute 2 visible";
                newSheet.Cells[$"S1"].Value = "Attribute 2 global";
                newSheet.Cells[$"T1"].Value = "Categories";
                newSheet.Cells[$"U1"].Value = "Allow customer reviews?";
                newSheet.Cells[$"V1"].Value = "Tags";
                newSheet.Cells[$"W1"].Value = "Images";
                newSheet.Cells[$"X1"].Value = "Parent";
                newSheet.Cells[$"Y1"].Value = "Regular price";
                
                //Abecedario Corto
                char[] alpha = "GHIJKLM".ToCharArray();
                int pointer = 0;

                //char[] alpha = "ABCDEFGHIJKLMNOPQRSTUVWXY".ToCharArray();
                //rawSheet.Dimension.Rows;
                for (int i = 2; i < rawSheet.Dimension.Rows; i++)
                {
                    Console.WriteLine($"Cell A{i} Value   : {rawSheet.Cells[$"A{i}"].Text}");
                    Console.WriteLine($"Cell B{i} Value   : {rawSheet.Cells[$"B{i}"].Text}");
                    Console.WriteLine($"Cell C{i} Value   : {rawSheet.Cells[$"C{i}"].Text}");
                    Console.WriteLine($"Cell D{i} Value   : {rawSheet.Cells[$"D{i}"].Text}");
                    Console.WriteLine($"Cell E{i} Value   : {rawSheet.Cells[$"E{i}"].Text}");
                    Console.WriteLine($"Cell F{i} Value   : {rawSheet.Cells[$"F{i}"].Text}");
                    Console.WriteLine($"Cell G{i} Value   : {rawSheet.Cells[$"G{i}"].Text}");
                    Console.WriteLine($"Cell H{i} Value   : {rawSheet.Cells[$"H{i}"].Text}");
                    Console.WriteLine($"Cell I{i} Value   : {rawSheet.Cells[$"I{i}"].Text}");
                    Console.WriteLine($"Cell J{i} Value   : {rawSheet.Cells[$"J{i}"].Text}");
                    Console.WriteLine($"Cell K{i} Value   : {rawSheet.Cells[$"K{i}"].Text}");
                    Console.WriteLine($"Cell L{i} Value   : {rawSheet.Cells[$"L{i}"].Text}");
                    Console.WriteLine($"Cell M{i} Value   : {rawSheet.Cells[$"M{i}"].Text}");
                    Console.WriteLine($"Cell N{i} Value   : {rawSheet.Cells[$"N{i}"].Text}");
                    Console.WriteLine($"Cell O{i} Value   : {rawSheet.Cells[$"O{i}"].Text}");
                    Console.WriteLine($"Cell P{i} Value   : {rawSheet.Cells[$"P{i}"].Text}");
                    Console.WriteLine($"Cell Q{i} Value   : {rawSheet.Cells[$"Q{i}"].Text}");

                    List<int> listChildren = new List<int>();
                    List<string> listChildrenStock = new List<string>();

                    foreach(var letter in alpha)
                    {
                        if (int.Parse(rawSheet.Cells[$"{letter}{i}"].Text) > 0)
                        {
                            listChildren.Add(int.Parse(rawSheet.Cells[$"{letter}1"].Text));
                            listChildrenStock.Add(rawSheet.Cells[$"G{i}"].Text);
                        }
                    }

                    # region Children If Manual
                    /*
                    if (int.Parse(rawSheet.Cells[$"G{i}"].Text) > 0)
                    {
                        listChildren.Add(5);
                        listChildrenStock.Add(rawSheet.Cells[$"G{i}"].Text);
                    }

                    if (int.Parse(rawSheet.Cells[$"H{i}"].Text) > 0)
                    {
                        listChildren.Add(6);
                        listChildrenStock.Add(rawSheet.Cells[$"H{i}"].Text);
                    }

                    if (int.Parse(rawSheet.Cells[$"I{i}"].Text) > 0)
                    {
                        listChildren.Add(7);
                        listChildrenStock.Add(rawSheet.Cells[$"I{i}"].Text);

                    }

                    if (int.Parse(rawSheet.Cells[$"J{i}"].Text) > 0)
                    {
                        listChildren.Add(8);
                        listChildrenStock.Add(rawSheet.Cells[$"J{i}"].Text);
                    }

                    if (int.Parse(rawSheet.Cells[$"K{i}"].Text) > 0)
                    {
                        listChildren.Add(9);
                        listChildrenStock.Add(rawSheet.Cells[$"K{i}"].Text);
                    }

                    if (int.Parse(rawSheet.Cells[$"L{i}"].Text) > 0)
                    {
                        listChildren.Add(10);
                        listChildrenStock.Add(rawSheet.Cells[$"L{i}"].Text);
                    }

                    if (int.Parse(rawSheet.Cells[$"M{i}"].Text) > 0)
                    {
                        listChildren.Add(11);
                        listChildrenStock.Add(rawSheet.Cells[$"M{i}"].Text);
                    }
                    */
                    #endregion

                    string sku = rawSheet.Cells[$"B{i}"].Text.Replace(".", "-");

                    //Parent
                    newSheet.Cells[$"A{i + pointer}"].Value = "variable";
                    newSheet.Cells[$"B{i + pointer}"].Value = sku;
                    newSheet.Cells[$"C{i + pointer}"].Value = rawSheet.Cells[$"A{i}"].Text;
                    newSheet.Cells[$"D{i + pointer}"].Value = "1";
                    newSheet.Cells[$"E{i + pointer}"].Value = "0";
                    newSheet.Cells[$"F{i + pointer}"].Value = "visible";
                    newSheet.Cells[$"G{i + pointer}"].Value = rawSheet.Cells[$"F{i}"].Text;
                    newSheet.Cells[$"H{i + pointer}"].Value = rawSheet.Cells[$"F{i}"].Text;
                    newSheet.Cells[$"I{i + pointer}"].Value = "1";
                    newSheet.Cells[$"J{i + pointer}"].Value = "";
                    newSheet.Cells[$"K{i + pointer}"].Value = "";
                    newSheet.Cells[$"L{i + pointer}"].Value = "Color";
                    newSheet.Cells[$"M{i + pointer}"].Value = rawSheet.Cells[$"C{i}"].Text;
                    newSheet.Cells[$"N{i + pointer}"].Value = "1";
                    newSheet.Cells[$"O{i + pointer}"].Value = "0";
                    newSheet.Cells[$"P{i + pointer}"].Value = "Tamaño";
                    newSheet.Cells[$"Q{i + pointer}"].Value = string.Join(",", listChildren);
                    newSheet.Cells[$"R{i + pointer}"].Value = "1";
                    newSheet.Cells[$"S{i + pointer}"].Value = "0";
                    newSheet.Cells[$"T{i + pointer}"].Value = rawSheet.Cells[$"E{i}"].Text;
                    newSheet.Cells[$"U{i + pointer}"].Value = "1";
                    newSheet.Cells[$"V{i + pointer}"].Value = "Precio de venta es el monto que se muestra tachado.";
                    newSheet.Cells[$"W{i + pointer}"].Value = $"https://mybusinessdropshipping.com/wp-content/uploads/2020/12/{sku}.jpg";
                    newSheet.Cells[$"X{i + pointer}"].Value = "";
                    newSheet.Cells[$"Y{i + pointer}"].Value = rawSheet.Cells[$"N{i}"].Text; ;

                    for (int x = 0; x < listChildren.Count; x++)
                    {
                        pointer++;
                        newSheet.Cells[$"A{i+pointer}"].Value = "variation";
                        newSheet.Cells[$"B{i+pointer}"].Value = $"{sku}-{x+1}";
                        newSheet.Cells[$"C{i+pointer}"].Value = rawSheet.Cells[$"A{i}"].Text;
                        newSheet.Cells[$"D{i+pointer}"].Value = "1";
                        newSheet.Cells[$"E{i+pointer}"].Value = "0";
                        newSheet.Cells[$"F{i+pointer}"].Value = "visible";
                        newSheet.Cells[$"G{i+pointer}"].Value = "";
                        newSheet.Cells[$"H{i+pointer}"].Value = "";
                        newSheet.Cells[$"I{i+pointer}"].Value = "1";
                        newSheet.Cells[$"J{i+pointer}"].Value = listChildrenStock[x];
                        newSheet.Cells[$"K{i+pointer}"].Value = $"{x + 1}";
                        newSheet.Cells[$"L{i+pointer}"].Value = "Tamaño";
                        newSheet.Cells[$"M{i+pointer}"].Value = listChildren[x];
                        newSheet.Cells[$"N{i+pointer}"].Value = "1";
                        newSheet.Cells[$"O{i+pointer}"].Value = "0";
                        newSheet.Cells[$"P{i+pointer}"].Value = "";
                        newSheet.Cells[$"Q{i+pointer}"].Value = "";
                        newSheet.Cells[$"R{i+pointer}"].Value = "";
                        newSheet.Cells[$"S{i+pointer}"].Value = "";
                        newSheet.Cells[$"T{i+pointer}"].Value = "";
                        newSheet.Cells[$"U{i+pointer}"].Value = "1";
                        newSheet.Cells[$"V{i+pointer}"].Value = "Precio de venta es el monto que se muestra tachado.";
                        newSheet.Cells[$"W{i+pointer}"].Value = "";
                        newSheet.Cells[$"X{i+pointer}"].Value = sku;
                        newSheet.Cells[$"Y{i+pointer}"].Value = rawSheet.Cells[$"N{i}"].Text;
                    }
                }
                package.Save();
            }
        }
    }
}
