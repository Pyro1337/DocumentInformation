using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using LibreOfficeOperations;

public class ConvertToPdf
{
    private readonly string directoryTemp;
    private readonly Log log;
    public static readonly List<string> officeExtensions = new List<string> { "docx", "doc", "xlsx", "xls", "pptx", "ppt" };

    public ConvertToPdf(string _directoryTemp)
    {
        directoryTemp = _directoryTemp;
        log = new Log($@"{directoryTemp}");
    }

    public string OfficeToPDF(byte[] fileBtytes, string extension, string pathSofficeExe)
    {
        string pathPdf = "";
        try
        {
            if (extension.Substring(0, 1).Equals(".")) //Obtiene desde el comienzo hasta el segundo caracter
            {
                extension = extension.Substring(1, extension.Length - 1);
            }
            if (officeExtensions.Contains(extension.ToLower()))//Si es de office
            {
                pathPdf = Convert(fileBtytes, extension, pathSofficeExe);
            }
            else
            {
                log.Add($"Extensión {extension} no valida para la conversión a pdf.");
            }
            return pathPdf;
        }
        catch (Exception e)
        {
            log.Add(e.Message);
            return "";
        }
    }

    public string OfficeToPDF(string pathFile, string pathSofficeExe)
    {
        string extension = Path.GetExtension(pathFile);
        string pathPdf = "";
        try
        {
            if (extension.Substring(0, 1).Equals(".")) //Obtiene desde el comienzo hasta el segundo caracter
            {
                extension = extension.Substring(1, extension.Length - 1);
            }
            if (officeExtensions.Contains(extension.ToLower()))//Si es de office
            {
                pathPdf = Convert(pathFile, pathSofficeExe);
            }
            else
            {
                log.Add($"Extensión {extension} no valida para la conversión a pdf.");
            }
            return pathPdf;
        }
        catch (Exception e)
        {
            log.Add(e.Message);
            return "";
        }
    }

    private string Convert(byte[] fileBtytes, string extension, string pathSofficeExe)
    {
        string pathFile = "";
        try
        {
            if (!File.Exists(pathSofficeExe) || !pathSofficeExe.Contains("soffice.exe"))
            {
                log.Add("Ruta de soffice.exe incorrecta");
                return "";
            }
            pathFile = SaveDisk(fileBtytes, extension);
            string pathResult = Path.Combine(Path.GetDirectoryName(pathFile), Guid.NewGuid().ToString() + ".pdf");
            DocumentConverter converter = new DocumentConverter(); //Librería https://github.com/DigDes/LibreOfficeLibrary/tree/master  
            converter.ConvertToPdf(pathFile, pathResult, Path.GetDirectoryName(pathSofficeExe));
            if (!File.Exists(pathResult)) { return ""; } //Se verifica que exista realmente el pdf
            DeleteFile(pathFile);//Se elimina el archivo office escrito en disco y se queda el pdf
            return pathResult;//Se retorna la ruta del pdf
        }
        catch (Exception) { throw; }
        finally { DeleteFile(pathFile); }
    }

    private string Convert(string pathFile, string pathSofficeExe)
    {
        try
        {
            if (!File.Exists(pathSofficeExe) || !pathSofficeExe.Contains("soffice.exe"))
            {
                log.Add("Ruta de soffice.exe incorrecta");
                return "";
            }
            string pathResult = Path.Combine(Path.GetDirectoryName(pathFile), Guid.NewGuid().ToString() + ".pdf");
            DocumentConverter converter = new DocumentConverter(); //Librería https://github.com/DigDes/LibreOfficeLibrary/tree/master  
            converter.ConvertToPdf(pathFile, pathResult, Path.GetDirectoryName(pathSofficeExe));
            if (!File.Exists(pathResult)) { return ""; } //Se verifica que exista realmente el pdf
            return pathResult;//Se retorna la ruta del pdf
        }
        catch(Exception) { throw; }
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
        CreateDirectory(pathFile);
        File.WriteAllBytes(pathFile, fileBtytes);
        return pathFile;
    }

    private void CreateDirectory(string path)
    {
        string directory = Path.GetDirectoryName(path);
        if (!Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    private void DeleteFile(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

}

