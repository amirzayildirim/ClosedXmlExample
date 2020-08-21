using ClosedXML.Excel;
using ExcelExportTraining.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExcelExportTraining.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View( GetCustomer() );
        }

        public List<Customer> GetCustomer()
        {
            List<Customer> CustomerList = new List<Customer>();

            for (int i = 1; i <= 50; i++)
            {
                CustomerList.Add
                    (
                        new Customer
                        {   ID = i,
                            Name = "Name - " + i,
                            Surname = "Surname - " + i,
                            Phone = i.ToString()
                        }
                    );
            }

            return CustomerList;
        }

        public FileStreamResult ExportToExcel(string id)
        {
            DataTable dt = new DataTable();
         
            dt.Columns.AddRange
           (
               new DataColumn[3]
               {
                            new DataColumn("Isim"),
                            new DataColumn("Soyisim"),
                            new DataColumn("Telefon")
               }
           );

            dt.TableName = "Müşteri Listesi"; // Excel içinde ki sayfada görüntülenecek isim
            List<Customer> CustomerList = GetCustomer();
            int rowIndex = 0;

            foreach (var Customer in CustomerList)
            {
                dt.Rows.Add();
                dt.Rows[rowIndex]["Isim"] = Customer.Name;
                dt.Rows[rowIndex]["Soyisim"] = Customer.Surname;
                dt.Rows[rowIndex]["Telefon"] = Customer.Phone;

                rowIndex++;
            }

            // dosya ismi hazırlanıyor . aykırı karakterler temizleniyor..
            string fileName = ("Müşteri Listesi-" + DateTime.Now.ToString()).Replace('/', '_').Replace(':', '.').Replace(' ', '_') + ".xlsx";


            using (var workbook = new XLWorkbook())
            {
                MemoryStream Ms = new MemoryStream();

                var ws = workbook.Worksheets.Add(dt);
                ws.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center); // öğeler yatayda ortalanıyor
                ws.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center); // öğeler dikeyde ortalanıyor
               
                /*
                 Satır ve sütunların boyutu içeriklere göre ayarlanıyor.
                 */
                ws.Rows().AdjustToContents();
                ws.Columns().AdjustToContents();
              
                /*
                 false => Bütün kolonlar için otomatik filtreleme kapatılıyor
                 true => Bütün kolonlar için otomatik filtreleme aktif oluyor(varsayılan değer)
                */
                ws.Tables.FirstOrDefault().ShowAutoFilter = false;

                workbook.SaveAs(Ms);
                Ms.Position = 0;

                // File stream olarak sonuç döndürülüyor.
                return new FileStreamResult(Ms, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") { FileDownloadName = fileName };
            }
        }
    }
}