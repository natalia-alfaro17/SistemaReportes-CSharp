
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using ClosedXML.Excel;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace reportes
{
    public class Program
    {
        static string connectionString = @"Server=(localdb)\MSSQLLocalDB;Database=UniversidadDB;Trusted_Connection=True;TrustServerCertificate=True;";

        public static void Main(string[] args)
        {
            QuestPDF.Settings.License = LicenseType.Community;

            while (true)
            {
                Console.Clear();
                Console.WriteLine("=============================================");
                Console.WriteLine(" UNIVERSIDAD DR. ANDRÉS BELLO - PROGRAMACIÓN II");
                Console.WriteLine("       SISTEMA DE REPORTES EMPRESARIALES");
                Console.WriteLine("=============================================");
                Console.WriteLine("1. Generar Reporte PDF");
                Console.WriteLine("2. Generar Reporte Excel");
                Console.WriteLine("3. Generar Ambos Reportes");
                Console.WriteLine("4. Salir");
                Console.Write("\nSeleccione una opción: ");

                string opcion = Console.ReadLine();

                switch (opcion)
                {
                    case "1":
                        GenerarPDF();
                        break;
                    case "2":
                        GenerarExcel();
                        break;
                    case "3":
                        GenerarPDF();
                        GenerarExcel();
                        break;
                    case "4":
                        return;
                    default:
                        Console.WriteLine("\nOpción no válida.");
                        Console.WriteLine("Presione Enter para continuar...");
                        Console.ReadLine();
                        break;
                }
            }
        }

        static DataTable ObtenerDatos(string query)
        {
            DataTable dt = new DataTable();

            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand(query, conn))
            {
                conn.Open();

                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    adapter.Fill(dt);
                }
            }

            return dt;
        }

        static DataTable ObtenerDatosCalculados()
        {
            string query = @"
                SELECT
                    COUNT(*) AS [Total Registros],
                    ISNULL(SUM(CAST(Monto AS DECIMAL(18,2))), 0) AS [Suma Total],
                    ISNULL(AVG(CAST(Monto AS DECIMAL(18,2))), 0) AS [Promedio]
                FROM PagosRealizados;";

            return ObtenerDatos(query);
        }

        static void GenerarExcel()
        {
            try
            {
                DataTable dtGeneral = ObtenerDatos("SELECT * FROM PagosRealizados;");
                DataTable dtConsulta = ObtenerDatosCalculados();

                using (XLWorkbook libro = new XLWorkbook())
                {
                    var hoja1 = libro.Worksheets.Add("Pagos Realizados");

                    hoja1.Cell("A1").Value = "REPORTE GENERAL DE PAGOS REALIZADOS";
                    hoja1.Cell("A1").Style.Font.Bold = true;
                    hoja1.Cell("A1").Style.Font.FontSize = 16;
                    hoja1.Range("A1:F1").Merge();

                    hoja1.Cell("A2").Value = "Fecha de generación:";
                    hoja1.Cell("B2").Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

                    var tabla1 = hoja1.Cell("A4").InsertTable(dtGeneral);
                    tabla1.Theme = XLTableTheme.TableStyleMedium2;

                    hoja1.Columns().AdjustToContents();
                    hoja1.SheetView.FreezeRows(4);

                    var hoja2 = libro.Worksheets.Add("Resumen Evaluado");

                    hoja2.Cell("A1").Value = "RESUMEN Y MÉTRICAS DE PAGOS";
                    hoja2.Cell("A1").Style.Font.Bold = true;
                    hoja2.Cell("A1").Style.Font.FontSize = 16;
                    hoja2.Range("A1:C1").Merge();

                    hoja2.Cell("A2").Value = "Fecha de generación:";
                    hoja2.Cell("B2").Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

                    var tabla2 = hoja2.Cell("A4").InsertTable(dtConsulta);
                    tabla2.Theme = XLTableTheme.TableStyleMedium2;

                    hoja2.Columns().AdjustToContents();

                    string path = Path.Combine(
                        AppDomain.CurrentDomain.BaseDirectory,
                        "Reportes_Universidad.xlsx"
                    );

                    libro.SaveAs(path);

                    Console.WriteLine("\n¡Excel generado exitosamente!");
                    Console.WriteLine("Archivo: " + path);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("\nError al generar Excel: " + ex.Message);
            }

            Console.WriteLine("\nPresione Enter para continuar...");
            Console.ReadLine();
        }

        static void GenerarPDF()
        {
            try
            {
                DataTable dtGeneral = ObtenerDatos("SELECT * FROM PagosRealizados;");
                DataTable dtConsulta = ObtenerDatosCalculados();

                string path = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    "Reportes_Universidad.pdf"
                );

                Document.Create(document =>
                {
                    document.Page(page =>
                    {
                        page.Size(PageSizes.A4);
                        page.Margin(30);

                        page.Header()
                            .Column(header =>
                            {
                                header.Item()
                                    .Text("UNIVERSIDAD DR. ANDRÉS BELLO")
                                    .FontSize(18)
                                    .Bold()
                                    .FontColor(Colors.Blue.Medium);

                                header.Item()
                                    .Text("SISTEMA DE REPORTES EMPRESARIALES")
                                    .FontSize(12)
                                    .Bold();

                                header.Item()
                                    .Text("Fecha de generación: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm"))
                                    .FontSize(9);
                            });

                        page.Content().Column(column =>
                        {
                            column.Spacing(15);

                            column.Item()
                                .Text("1. Listado General de Pagos Registrados")
                                .FontSize(14)
                                .Bold();

                            AgregarTablaQuestPDF(column, dtGeneral);

                            column.Item()
                                .Text("2. Resumen y Métricas Calculadas")
                                .FontSize(14)
                                .Bold();

                            AgregarTablaQuestPDF(column, dtConsulta);
                        });

                        page.Footer()
                            .AlignCenter()
                            .Text(text =>
                            {
                                text.Span("Reporte generado para Programación II | Página ");
                                text.CurrentPageNumber();
                                text.Span(" de ");
                                text.TotalPages();
                            });
                    });
                }).GeneratePdf(path);

                Console.WriteLine("\n¡PDF generado exitosamente!");
                Console.WriteLine("Archivo: " + path);
            }
            catch (Exception ex)
            {
                Console.WriteLine("\nError al generar PDF: " + ex.Message);
            }

            Console.WriteLine("\nPresione Enter para continuar...");
            Console.ReadLine();
        }

        static void AgregarTablaQuestPDF(ColumnDescriptor column, DataTable dt)
        {
            if (dt.Rows.Count == 0)
            {
                column.Item()
                    .Text("No hay registros disponibles.")
                    .Italic();

                return;
            }

            column.Item().Table(table =>
            {
                table.ColumnsDefinition(columns =>
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        columns.RelativeColumn();
                    }
                });

                table.Header(header =>
                {
                    foreach (DataColumn col in dt.Columns)
                    {
                        header.Cell()
                            .Background(Colors.Blue.Medium)
                            .Padding(5)
                            .Text(col.ColumnName)
                            .FontColor(Colors.White)
                            .Bold()
                            .FontSize(9);
                    }
                });

                foreach (DataRow row in dt.Rows)
                {
                    foreach (var item in row.ItemArray)
                    {
                        table.Cell()
                            .BorderBottom(1)
                            .BorderColor(Colors.Grey.Lighten3)
                            .Padding(5)
                            .Text(item?.ToString() ?? "")
                            .FontSize(9);
                    }
                }
            });
        }
    }
}
