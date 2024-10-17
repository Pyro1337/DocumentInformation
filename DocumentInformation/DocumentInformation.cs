using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;

/* Documentación de la librería LibreOfficeLibrary que se usa en la clase ConvertToPdf: https://github.com/DigDes/LibreOfficeLibrary/tree/master */
public class DocumentInformation
{
    private readonly string directoryTemp = $@"C:\DMSGX\tempLibreria";
    private readonly List<string> filesWithOnePage = new List<string> { "png", "jpg", "jpeg", "txt", "tiff", "html", "gif", "csv" };
    private readonly Log log = new Log($@"C:\DMSGX\tempLibreria");

    public int GetCantPagesFromBinaryFile(byte[] fileBytes, string extension, string pathSofficeExe)
    {
        try
        {
            int cantPages;
            if (extension.Substring(0, 1).Equals("."))//Por si la extensión viene con "." al inicio, se le saca
                extension = extension.Substring(1, extension.Length - 1);

            if (extension.ToLower().Equals("pdf"))
            {
                string pathFilePdf = SaveDisk(fileBytes, extension.ToLower());
                cantPages = GetCantPages(pathFilePdf);
                DeleteFile(pathFilePdf);
            }
            else if (ConvertToPdf.officeExtensions.Contains(extension.ToLower())) //Verificamos si es un archivo office
            {
                ConvertToPdf convertToPdf = new ConvertToPdf(directoryTemp);
                string pathFilePdf = convertToPdf.OfficeToPDF(fileBytes, extension, pathSofficeExe);//Retorna la ruta donde se generó el pdf
                if (!File.Exists(pathFilePdf)) { return 0; }
                cantPages = GetCantPages(pathFilePdf);
                DeleteFile(pathFilePdf);
            }
            else if (filesWithOnePage.Contains(extension.ToLower()))//Archivos que solo pueden tener una imagen. Ej: imagenes, txt, etc.
            {
                cantPages = 1;
            }
            else //Extensiones que no corresponden para el conteo de páginas, como audios, videos, etc.
            {
                cantPages = 0;
            }
            return cantPages;
        }
        catch (Exception e)
        {
            log.Add(e.Message);
            return 0;
        }
    }

    public int GetCantPagesFromPathFile(string pathFile, string pathSofficeExe)
    {
        try
        {
            string extension = Path.GetExtension(pathFile);
            int cantPages;
            if (extension.Substring(0, 1).Equals("."))//Por si la extensión viene con "." al inicio, se le saca
                extension = extension.Substring(1, extension.Length - 1);

            if (extension.ToLower().Equals("pdf"))
            {
                cantPages = GetCantPages(pathFile);
            }
            else if (ConvertToPdf.officeExtensions.Contains(extension.ToLower())) //Verificamos si es un archivo office
            {
                ConvertToPdf convertToPdf = new ConvertToPdf(directoryTemp);
                string pathFilePdf = convertToPdf.OfficeToPDF(pathFile, pathSofficeExe);//Retorna la ruta donde se generó el pdf
                if(!File.Exists(pathFilePdf)) { return 0; }
                cantPages = GetCantPages(pathFilePdf);
                DeleteFile(pathFilePdf);
            }
            else if (filesWithOnePage.Contains(extension.ToLower()))//Archivos que solo pueden tener una imagen. Ej: imagenes, txt, etc.
            {
                cantPages = 1;
            }
            else //Extensiones que no corresponden para el conteo de páginas, como audios, videos, etc.
            {
                cantPages = 0;
            }
            return cantPages;
        }
        catch (Exception e)
        {
            log.Add(e.Message);
            return 0;
        }
    }

    public int GetCantPagesFromBase64File(string fileBase64, string extension, string pathSofficeExe)
    {
        try
        {
            byte[] fileBytes = Convert.FromBase64String(fileBase64);
            int cantPages;
            string pathFilePdf = "";
            if (extension.Substring(0, 1).Equals("."))//Por si la extensión viene con "." al inicio, se le saca
                extension = extension.Substring(1, extension.Length - 1);

            if (extension.ToLower().Equals("pdf"))
            {
                pathFilePdf = SaveDisk(fileBytes, extension.ToLower());
                cantPages = GetCantPages(pathFilePdf);
                DeleteFile(pathFilePdf);
            }
            else if (ConvertToPdf.officeExtensions.Contains(extension.ToLower())) //Verificamos si es un archivo office
            {
                ConvertToPdf convertToPdf = new ConvertToPdf(directoryTemp);
                pathFilePdf = convertToPdf.OfficeToPDF(fileBytes, extension, pathSofficeExe);//Retorna la ruta donde se generó el pdf
                if (!File.Exists(pathFilePdf)) { return 0; }
                cantPages = GetCantPages(pathFilePdf);
                DeleteFile(pathFilePdf);
            }
            else if (filesWithOnePage.Contains(extension.ToLower()))//Archivos que solo pueden tener una imagen. Ej: imagenes, txt, etc.
            {
                cantPages = 1;
            }
            else //Extensiones que no corresponden para el conteo de páginas, como audios, videos, etc.
            {
                cantPages = 0;
            }
            return cantPages;
        }
        catch (Exception e)
        {
            log.Add(e.Message);
            return 0;
        }
    }

    public decimal GetFileSizeBytes(byte[] fileBytes, string extension)
    {
        try 
        {
            string pathFile = SaveDisk(fileBytes, extension);
            FileInfo fileInfo = new FileInfo(pathFile);
            string sizeKbAux = fileInfo.Length.ToString(); //Retorna el tamaño en bytes
            decimal sizeKB = decimal.Parse(sizeKbAux) / 1024; //Pasamos a kb
            DeleteFile(pathFile); //Se elimina el archivo temporal
            return sizeKB;
        }
        catch (Exception e)
        {
            log.Add("Error en el método GetFileSize(): " + e.Message);
            return 0;
        }
    }

    public decimal GetFileSizePath(string pathFile)
    {
        try
        {
            string extension = Path.GetExtension(pathFile);
            FileInfo fileInfo = new FileInfo(pathFile);
            string sizeKbAux = fileInfo.Length.ToString(); //Retorna el tamaño en bytes
            decimal sizeKB = decimal.Parse(sizeKbAux) / 1024; //Pasamos a kb
            return sizeKB;
        }
        catch (Exception e)
        {
            log.Add("Error en el método GetFileSize(): " + e.Message);
            return 0;
        }
    }

    private int GetCantPages(string pathFilePdf)
    {
        try
        {
            PdfReader reader = new PdfReader(pathFilePdf);
            int cantPages = reader.NumberOfPages;
            reader.Close();
            return cantPages;
        }
        catch (Exception e)
        {
            log.Add("Error en el método GetCantPages(): " + e.Message);
            return 0;
        }
    }

    private string SaveDisk(byte[] fileBtytes, string extension)
    {
        string nameFile = Guid.NewGuid().ToString();
        if (extension.Substring(0, 1).Equals(".")) //Obtiene desde el comienzo hasta el segundo caracter
        {
            extension = extension.Substring(1, extension.Length - 1); //Obtiene desde el segundo caracter hasta el final
        }
        extension = extension.ToLower();
        string pathFile = Path.Combine(directoryTemp, $"{nameFile}.{extension}");
        CreateDirectory(directoryTemp);
        File.WriteAllBytes(pathFile, fileBtytes);
        return pathFile;
    }

    private void CreateDirectory(string directory)
    {
        if (!Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    private void DeleteFile(string pathFile)
    {
        if (File.Exists(pathFile))
        {
            File.Delete(pathFile);
        }
    }
}
