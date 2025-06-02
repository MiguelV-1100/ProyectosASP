using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace SepararPagPDF
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Ruta específica para el archivo PDF
            string defaultInputPdfPath = @"C:\Users\Miguel\Desktop\Editora Perú\Envío CTS\DocCTS\CTS-MAYO-2025-1-50.pdf";

            try
            {
                if (File.Exists(defaultInputPdfPath))
                {
                    SplitPdfIntoPages(defaultInputPdfPath);
                    Console.WriteLine("¡Documento dividido con éxito!");
                    Console.WriteLine($"Páginas guardadas en: {System.IO.Path.GetDirectoryName(defaultInputPdfPath)}");
                }
                else
                {
                    Console.WriteLine("ERROR: No se encontró el archivo PDF en la ruta especificada.");
                    Console.WriteLine($"Ruta buscada: {defaultInputPdfPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error inesperado: " + ex.Message);
            }


            // Ruta de la carpeta con los PDFs divididos
            string pdfsFolderPath = @"C:\Users\Miguel\Desktop\Editora Perú\Envío CTS\DocCTS";

            if (!Directory.Exists(pdfsFolderPath))
            {
                Console.WriteLine($"Error: No se encontró la carpeta: {pdfsFolderPath}");
                Console.ReadKey();
                return;
            }

            string[] pdfFiles = Directory.GetFiles(pdfsFolderPath, "CTS-MAYO-2025-1-50_Pagina_*.pdf");

            if (pdfFiles.Length == 0)
            {
                Console.WriteLine("No se encontraron archivos PDF en la carpeta.");
                Console.ReadKey();
                return;
            }

            Console.WriteLine($"Procesando {pdfFiles.Length} archivos PDF...\n");

            foreach (string pdfFile in pdfFiles)
            {
                string fileName = System.IO.Path.GetFileName(pdfFile);
                string matricula = ExtractMatriculaFromPdf(pdfFile);

                if (!string.IsNullOrEmpty(matricula))
                {
                    // Nuevo nombre del archivo
                    string newFileName = $"{matricula}.pdf";
                    string newFilePath = System.IO.Path.Combine(pdfsFolderPath, newFileName);

                    // Renombrar el archivo
                    try
                    {
                        File.Move(pdfFile, newFilePath);
                        Console.WriteLine($"Renombrado: {fileName} -> {newFileName}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error al renombrar {fileName}: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"{fileName}: No se encontró matrícula (no se renombrará).");
                }
            }

            // Crear subcarpeta "CTS" si no existe
            string ctsFolderPath = System.IO.Path.Combine(pdfsFolderPath, "CTS");
            Directory.CreateDirectory(ctsFolderPath); // No hace nada si ya existe

            // Buscar todos los PDFs que sigan el patrón "DocCTS_Página_*.pdf"
            string[] archivosPDF = Directory.GetFiles(pdfsFolderPath, "000000????.pdf");

            if (archivosPDF.Length == 0)
            {
                Console.WriteLine("No se encontraron archivos PDF en la carpeta.");
                Console.ReadKey();
                return;
            }

            Console.WriteLine($"🔍 Procesando {archivosPDF.Length} archivos PDF...\n");

            foreach (string pdfFile in archivosPDF)
            {
                string fileName = System.IO.Path.GetFileName(pdfFile);
                string matricula = ExtractMatriculaFromPdf(pdfFile);

                if (!string.IsNullOrEmpty(matricula))
                {
                    string newFileName = $"{matricula}.pdf";
                    string newFilePath = System.IO.Path.Combine(ctsFolderPath, newFileName);

                    try
                    {
                        // Mover (no copiar) el archivo a la subcarpeta "CTS"
                        File.Move(pdfFile, newFilePath);
                        Console.WriteLine($"Movido: {fileName} -> /CTS/{newFileName}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error al mover {fileName}: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"{fileName}: No se encontró matrícula (no se movió).");
                }
            }

            Console.WriteLine("\nProceso completado...");
            Console.ReadKey();

        }


        static void SplitPdfIntoPages(string inputPdfPath)
        {
            // Ruta específica para el archivo PDF
            string outputDirectory = System.IO.Path.GetDirectoryName(inputPdfPath);
            string pdfFileName = System.IO.Path.GetFileNameWithoutExtension(inputPdfPath);

            using (PdfReader reader = new PdfReader(inputPdfPath))
            {
                for (int pageNum = 1; pageNum <= reader.NumberOfPages; pageNum++)
                {
                    using (Document document = new Document())
                    {
                        string outputPdfPath = System.IO.Path.Combine(
                            outputDirectory,
                            $"{pdfFileName}_Pagina_{pageNum}.pdf"  // Nombre personalizado
                        );

                        using (FileStream fs = new FileStream(outputPdfPath, FileMode.Create))
                        using (PdfCopy copy = new PdfCopy(document, fs))
                        {
                            document.Open();
                            copy.AddPage(copy.GetImportedPage(reader, pageNum));
                        }
                    }
                }
            }
        }

        static string ExtractMatriculaFromPdf(string pdfPath)
        {
            StringBuilder text = new StringBuilder();
            using (PdfReader reader = new PdfReader(pdfPath))
            {
                for (int page = 1; page <= reader.NumberOfPages; page++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string currentText = PdfTextExtractor.GetTextFromPage(reader, page, strategy);
                    text.Append(currentText);
                }
            }

            Regex regex = new Regex(@"\b\d{10}\b");
            Match match = regex.Match(text.ToString());

            return match.Success ? match.Value : null;
        }


    }
}
