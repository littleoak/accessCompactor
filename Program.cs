using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

public class AccessDBCompactor
{
    /// <summary>
    /// Compacta um banco de dados Access 97
    /// </summary>
    /// <param name="mdwfilename">Caminho completo do arquivo MDB que você deseja compactar.</param>
    public static void CompactAccessDB(string mdwfilename)
    {
        try
        {
            Type jetEngineType = Type.GetTypeFromProgID("JRO.JetEngine");
            if (jetEngineType == null)
            {
                throw new Exception("JRO.JetEngine não está registrado no sistema.");
            }

            object objJRO = Activator.CreateInstance(jetEngineType);

            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdwfilename;
            object[] oParams = new object[] {
                connectionString,
                "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\tempdb.mdb;Jet OLEDB:Engine Type=4"
            };

            objJRO.GetType().InvokeMember("CompactDatabase",
                BindingFlags.InvokeMethod,
                null,
                objJRO,
                oParams);

            File.Delete(mdwfilename);
            File.Move("C:\\tempdb.mdb", mdwfilename);

            Marshal.ReleaseComObject(objJRO);
            objJRO = null;
        }
        catch (COMException comEx)
        {
            Console.WriteLine("Erro durante a compactação: " + comEx.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Erro durante a compactação: " + ex.Message);
        }
    }

    public static void Main(string[] args)
    {
        if (args.Length != 1)
        {
            Console.WriteLine("Uso: AccessDBCompactor.exe <caminho_do_banco_de_dados>");
            return;
        }

        string dbPath = args[0];

        try
        {
            CompactAccessDB(dbPath);
            Console.WriteLine("Compactação concluída com sucesso.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Erro durante a compactação: " + ex.Message);
        }
    }
}
