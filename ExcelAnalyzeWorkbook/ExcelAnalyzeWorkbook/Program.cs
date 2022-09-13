using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAnalyzeWorkbook
{
    class FindFormatTemplate
    {
        public int interiorColorFilter;
        public int fontSizeFilter;

        private List<Object> comObjects;

        public void set(Excel.CellFormat cellFormat)
        {
            cellFormat.Clear();

            var interior = cellFormat.Interior;
            interior.Color = interiorColorFilter;

            var font = cellFormat.Font;
            font.Size = fontSizeFilter;

            comObjects = new List<object>
            {
                interior,
                font
            };

            foreach (object ob in comObjects)
            {
                Marshal.ReleaseComObject(ob);
            }
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Необходимо выбрать файл Excel с выгрузкой Тирика для корректировки. Нажмите любую клавишу для продолжения.");
            Console.ReadKey();

            
            CorrOfReport corr = new CorrOfReport();
            corr.Correction();
        }
    }

    public class CorrOfReport
    {
        private Excel.Application excelApp;
        private Excel.Workbooks workbooks;
        private Excel.Workbook wrkbk;
        private Excel._Worksheet wrksh;
        private Excel.Interior cfI;
        private Excel.Range cells;

        private List<Object> comObjects;

        string fileName;

        const int clName = 3;
        const int clCount = 18;
        const int clPrice = 21;
        const int clPriceSource = 30;

        const int blue = 0xFF0000;
        const int gray = 0xC0C0C0;
        const int white = 0xFFFFFF;

        private List<OrderRow> order;
        private void promptSelectFile()
        {
            dynamic response;
            response = excelApp.GetOpenFilename("Excel Files (*.xlsx;*.xls), *.xlsx;*.xls");
            if (Convert.ToString(response) == "False")
                fileName = "";
            else
                fileName = response;
        }

        private void OrderLinesToList()
        {
            cells = wrksh.Cells;
            var columns = wrksh.Columns;
            var generalColumn = columns[clName];

            var findFormat = excelApp.FindFormat;
            FindFormatTemplate tablix = new FindFormatTemplate
            {
                fontSizeFilter = 10,
                interiorColorFilter = white
            };
            tablix.set(findFormat);
            string searhText = "*";

            List<int> orderRowsNumber = new List<int>();

            var cellFinded = generalColumn.Find(searhText, SearchFormat: true);
            if (cellFinded is not null)
            {
                var firstResultAddress = cellFinded.Address;
                var cellFindedAddress = cellFinded.Address;
                do
                {
                    //Console.WriteLine("{0},{1}", cellFinded.Row, cellFinded.Column);
                    orderRowsNumber.Add(cellFinded.Row);
                    cfI = cellFinded.Interior;
                    var prevFoundCell = cellFinded;
                    cellFinded = generalColumn.Find(searhText, SearchFormat: true, After: prevFoundCell);
                    cellFindedAddress = cellFinded.Address;

                    Marshal.ReleaseComObject(cfI);
                    Marshal.ReleaseComObject(prevFoundCell);
                }
                while (cellFindedAddress != firstResultAddress);
            }
            else
            {
                cellFinded = columns.Find("");
            }

            if (orderRowsNumber.Count == 0)
            {
                Console.WriteLine("Не найдены строки для корректировки!");
                comObjects.AddRange(new List<object>
                {
                     columns
                    ,generalColumn
                    ,findFormat
                    ,cellFinded
                    ,cells
                });
                return;
            }

            order = new List<OrderRow>();
            int numberRow = 0;
            var ci = new System.Globalization.CultureInfo("ru-ru");
            try
            {
                foreach (int number in orderRowsNumber)
                {
                    numberRow = number;
                    var nameVal = cells[number, clName];
                    var priceVal = cells[number, clPrice];
                    var priceSourceVal = cells[number, clPriceSource];
                    var count = cells[number, clCount];

                    order.Add(new OrderRow
                    {
                        item = new Item { name = nameVal.Value },
                        price = new Price { value = Convert.ToDouble(priceVal.Value, ci) },
                        priceSource = new Price { value = Convert.ToDouble(priceSourceVal.Value, ci) },
                        count = Convert.ToUInt32(count.value)
                    });

                    Marshal.ReleaseComObject(nameVal);
                    Marshal.ReleaseComObject(priceVal);
                    Marshal.ReleaseComObject(priceSourceVal);
                    Marshal.ReleaseComObject(count);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Что-то пошло не так при считывании данных на строке " + numberRow);
            }
            finally
            {
                comObjects.AddRange(new List<object>
                {
                     columns
                    ,generalColumn
                    ,findFormat
                    ,cellFinded
                    ,cells
                });
            }
        }
        public void Correction()
        {
            try
            {

                excelApp = new Excel.Application();
                excelApp.Visible = true;

                promptSelectFile();
                //fileName = "C:\\byAlexey\\a.XLSX";
                if (!File.Exists(fileName))
                {
                    excelApp.Quit();
                    comObjects.AddRange(new List<object>
                    {
                        excelApp
                    });
                    Console.Clear();
                    Console.WriteLine("Выбор файла отменен.");
                    Console.WriteLine("Нажмите любую клавишу для завершения.");
                    Console.ReadKey();
                    return;
                }

                Console.WriteLine(fileName);
                workbooks = excelApp.Workbooks;
                wrkbk = workbooks.Open(fileName);
                wrksh = (Excel.Worksheet)excelApp.ActiveSheet;

                OrderLinesToList();
                wrkbk.Close(false);

                var corrBook = workbooks.Add();
                var corrWrksh = (Excel.Worksheet) excelApp.ActiveSheet;
                cells = corrWrksh.Cells;


                uint orderNumber = 4;

                cells[orderNumber, "D"] = "Реализация";
                cells[orderNumber, "F"] = "Закупка";
                var titlePrice = corrWrksh.Range["D4:E4"];
                var titlePriceSource = corrWrksh.Range["F4:G4"];
                titlePrice.Merge();
                titlePriceSource.Merge();

                orderNumber = 5;

                cells[orderNumber, "B"] = "Наименование";
                cells[orderNumber, "C"] = "К-во";
                cells[orderNumber, "D"] = "Цена";
                cells[orderNumber, "E"] = "Сумма";
                cells[orderNumber, "F"] = "Цена";
                cells[orderNumber, "G"] = "Сумма";

                orderNumber = 6;
                foreach (OrderRow orderRow in order)
                {
                    cells[orderNumber, "B"] = orderRow.item.name;
                    cells[orderNumber, "C"] = orderRow.count;
                    cells[orderNumber, "D"] = orderRow.price.value;
                    cells[orderNumber, "E"] = orderRow.amount;
                    cells[orderNumber, "F"] = orderRow.priceSource;
                    cells[orderNumber, "G"] = orderRow.amountBySource;

                    orderNumber++;
                }

                var corrColumns = corrWrksh.Columns["A:G"];
                corrColumns.AutoFit();
                
                corrBook.SaveAs(fileName.Insert(fileName.Length - 5, "-corr"));
                excelApp.Quit();


                comObjects.AddRange(new List<object>
                {
                     excelApp
                    ,workbooks
                    ,wrkbk
                    ,wrksh
                    ,corrBook
                    ,corrWrksh
                    ,corrColumns
                });

                Console.WriteLine("Корректировка успешно создана!");
                Console.WriteLine("Введите 'r' для повтора или нажмите любую другую клавишу для завершения.");
                char userResponse = (char) Console.ReadKey().Key;
                if (userResponse == 'R')
                {
                    Console.Clear();
                    foreach (object ob in comObjects)
                    {
                        Marshal.ReleaseComObject(ob);
                    }
                    this.Correction();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                Console.ReadKey();
            }

            finally
            {
                foreach (object ob in comObjects)
                {
                    Marshal.ReleaseComObject(ob);
                }
            }
        }
        public CorrOfReport()
        {
            comObjects = new List<object>();
        }
    }
    public class Item
    {
        public string name { get; set; }
    }
    public class Price
    {
        public double value { get; set; }
    }
    public class OrderRow
    {
        public Item item { get; set; }
        public Price price { get; set; }
        public Price priceSource { get; set; }
        public uint count { get; set; }

        public double amount
        {
            get
            {
                return price.value * count;
            }
        }
        public double amountBySource
        {
            get
            {
                return priceSource.value * count;
            }
        }
    }
}