using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace ManipulacionCTS
{
    internal class Program
    {
        // Configuración de rutas
        private const string RutaPdfInicial = @"C:\Users\Miguel\Desktop\Editora Perú\Envío CTS\DocCTS\CTS-MAYO-2025-1-50.pdf";
        private const string RutaCarpetaPdf = @"C:\Users\Miguel\Desktop\Editora Perú\Envío CTS\DocCTS";
        private const string NuevaSubCarpeta = "CTS";

        static void Main(string[] args)
        {
            try
            {
                EjecutarProceso();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error inesperado: {ex.Message}");
            }
            finally
            {
                Console.WriteLine("\nProceso completado. Presione cualquier tecla para salir...");
                Console.ReadKey();
            }
        }

        private static void EjecutarProceso()
        {
            // Paso 1: Dividir el PDF principal en páginas individuales
            if (DividirPaginas(RutaPdfInicial))
            {
                Console.WriteLine("¡Documento dividido con éxito!");
                Console.WriteLine($"Páginas guardadas en: {System.IO.Path.GetDirectoryName(RutaPdfInicial)}");
            }

            // Paso 2: Procesar las páginas divididas para extraer matrículas
            RenombrarPdfMasivo();

            // Paso 3: Organizar los PDFs con matrícula en subcarpeta
            MoverPdfMasivo();
        }

        private static bool DividirPaginas(string rutaPdf)
        {
            if (!File.Exists(rutaPdf))
            {
                Console.WriteLine($"ERROR: No se encontró el archivo PDF en la ruta especificada:\n{rutaPdf}");
                return false;
            }

            string CarpetaSalida = System.IO.Path.GetDirectoryName(rutaPdf);
            string NombrePdf = System.IO.Path.GetFileNameWithoutExtension(rutaPdf);

            try
            {
                using (PdfReader reader = new PdfReader(rutaPdf))
                {
                    for (int NumPagina = 1; NumPagina <= reader.NumberOfPages; NumPagina++)
                    {
                        using (Document document = new Document())
                        {
                            string outputPdfPath = System.IO.Path.Combine(
                                CarpetaSalida,
                                $"{NombrePdf}_Pagina_{NumPagina}.pdf"
                            );

                            using (FileStream fs = new FileStream(outputPdfPath, FileMode.Create))
                            using (PdfCopy copy = new PdfCopy(document, fs))
                            {
                                document.Open();
                                copy.AddPage(copy.GetImportedPage(reader, NumPagina));
                            }
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al dividir el PDF: {ex.Message}");
                return false;
            }
        }

        static string ExtraerMatricula(string rutaPdf)
        {
            StringBuilder text = new StringBuilder();
            using (PdfReader reader = new PdfReader(rutaPdf))
            {
                for (int pagina = 1; pagina <= reader.NumberOfPages; pagina++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string currentText = PdfTextExtractor.GetTextFromPage(reader, pagina, strategy);
                    text.Append(currentText);
                }
            }

            Regex regex = new Regex(@"\b\d{10}\b");
            Match match = regex.Match(text.ToString());

            return match.Success ? match.Value : null;
        }

        private static void RenombrarPdf(string archivoPdf)
        {
            string nombrePdf = System.IO.Path.GetFileName(archivoPdf);
            string matricula = ExtraerMatricula(archivoPdf);

            if (string.IsNullOrEmpty(matricula))
            {
                Console.WriteLine($"{nombrePdf}: No se encontró matrícula (no se renombrará).");
                return;
            }

            string nuevoNombrePdf = $"{matricula}.pdf";
            string nuevaRutaPdf = System.IO.Path.Combine(RutaCarpetaPdf, nuevoNombrePdf);

            try
            {
                File.Move(archivoPdf, nuevaRutaPdf);
                Console.WriteLine($"Renombrado: {nombrePdf} -> {nuevoNombrePdf}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al renombrar {nombrePdf}: {ex.Message}");
            }
        }

        private static void RenombrarPdfMasivo()
        {
            if (!Directory.Exists(RutaCarpetaPdf))
            {
                Console.WriteLine($"Error: No se encontró la carpeta: {RutaCarpetaPdf}");
                return;
            }

            var archivosPdf = Directory.GetFiles(RutaCarpetaPdf, "CTS-MAYO-2025-1-50_Pagina_*.pdf");
            if (archivosPdf.Length == 0)
            {
                Console.WriteLine("No se encontraron archivos PDF en la carpeta.");
                return;
            }

            Console.WriteLine($"\nProcesando {archivosPdf.Length} archivos PDF para extraer matrículas...");

            foreach (var archivoPdf in archivosPdf)
            {
                RenombrarPdf(archivoPdf);
            }
        }

        private static void MoverPdf(string rutaArchivoPdf, string RutaCarpeta)
        {
            string rutaPdf = System.IO.Path.GetFileName(rutaArchivoPdf);
            string matricula = ExtraerMatricula(rutaArchivoPdf);

            if (string.IsNullOrEmpty(matricula))
            {
                Console.WriteLine($"{rutaPdf}: No se encontró matrícula (no se moverá).");
                return;
            }

            string nuevoNombrePdf = $"{matricula}.pdf";
            string nuevaRutaPdf = System.IO.Path.Combine(RutaCarpeta, nuevoNombrePdf);

            try
            {
                File.Move(rutaArchivoPdf, nuevaRutaPdf);
                Console.WriteLine($"Movido: {rutaPdf} -> /{NuevaSubCarpeta}/{nuevoNombrePdf}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al mover {rutaPdf}: {ex.Message}");
            }
        }

        private static void MoverPdfMasivo()
        {
            var rutaCarpetaCTS = System.IO.Path.Combine(RutaCarpetaPdf, NuevaSubCarpeta);
            Directory.CreateDirectory(rutaCarpetaCTS);

            var archivosPDF = Directory.GetFiles(RutaCarpetaPdf, "00000?????.pdf");
            if (archivosPDF.Length == 0)
            {
                Console.WriteLine("\nNo se encontraron archivos PDF con matrículas para organizar.");
                return;
            }

            Console.WriteLine($"\nOrganizando {archivosPDF.Length} archivos PDF en la subcarpeta {NuevaSubCarpeta}...");

            foreach (var archivoPdf in archivosPDF)
            {
                MoverPdf(archivoPdf, rutaCarpetaCTS);
            }
        }

    }
}