using System;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.ServiceProcess;
using Microsoft.VisualBasic.FileIO;


class IniFile   // revision 11
{
    string Path;
    string EXE = Assembly.GetExecutingAssembly().GetName().Name;

    [DllImport("kernel32", CharSet = CharSet.Unicode)]
    static extern long WritePrivateProfileString(string Section, string Key, string Value, string FilePath);

    [DllImport("kernel32", CharSet = CharSet.Unicode)]
    static extern int GetPrivateProfileString(string Section, string Key, string Default, StringBuilder RetVal, int Size, string FilePath);

    public IniFile(string IniPath = null)
    {
        Path = new FileInfo(IniPath ?? EXE + ".ini").FullName;
    }

    public string Read(string Key, string Section = null)
    {
        var RetVal = new StringBuilder(255);
        GetPrivateProfileString(Section ?? EXE, Key, "", RetVal, 255, Path);
        return RetVal.ToString();
    }

    public void Write(string Key, string Value, string Section = null)
    {
        WritePrivateProfileString(Section ?? EXE, Key, Value, Path);
    }

    public void DeleteKey(string Key, string Section = null)
    {
        Write(Key, null, Section ?? EXE);
    }

    public void DeleteSection(string Section = null)
    {
        Write(null, null, Section ?? EXE);
    }

    public bool KeyExists(string Key, string Section = null)
    {
        return Read(Key, Section).Length > 0;
    }
}


class Folder
{
    public void CreateIfNotExists(string dir)
    {
        // Verificar se a pasta de destino existe
        if (!Directory.Exists(dir))
        {
            // Criar a pasta de destino se não existir
            Directory.CreateDirectory(dir);
        }
    }

    public void CopyFolder(string dirOrigin, string dirDestiny)
    {

        CreateIfNotExists(dirDestiny);

        // Copiar todos os arquivos da pasta de origem para a pasta de destino
        foreach (string file in Directory.GetFiles(dirOrigin))
        {
            string fileDestiny = Path.Combine(dirDestiny, Path.GetFileName(file));
            File.Copy(file, fileDestiny, true);
        }

        // Copiar todas as subpastas da pasta de origem para a pasta de destino
        foreach (string subfolder in Directory.GetDirectories(dirOrigin))
        {
            string subpastaDestino = Path.Combine(dirDestiny, Path.GetFileName(subfolder));
            CopyFolder(subfolder, subpastaDestino);
        }
    }   
    
    public void RenameFile(string dirFile, string actualFileName, string newFileName)
    {
        string oldName = dirFile + '\\' + actualFileName;
        string newName = dirFile + '\\' + newFileName;
        
        if (oldName != newName)
        {
            File.Copy(oldName, newName, true);
            File.Delete(oldName);
        }
    }
}
class Menu
{
    public ConfigIniFile configIniFile;
    public WindowsServicesManager windowsServicesManager;


    public Menu()
    {
        configIniFile = new ConfigIniFile("AtualizadorServicosProtheus.ini");
        windowsServicesManager = new WindowsServicesManager();

        // Inicia com o carregamento das variaveis
        configIniFile.LoadIniFile();

        // Ja inicia carregando a lista dos servicos
        windowsServicesManager.WriteListOfServices(configIniFile.GetSearchPatternList(), configIniFile.GetExcludePatternList());

        Console.Clear();

    }

    public void LoadIniFile()
    {
        Console.WriteLine("ATUALIZADOR DE SERVICOS PROTHEUS");
        Console.WriteLine("--------------------------------");
        Console.WriteLine("Duvidas: danilo.david@totvs.com.br");
        Console.WriteLine("----------------------------------");
        Console.WriteLine("");
        Console.WriteLine("Parametros definidos no arquivo de configuracoes");
        Console.WriteLine("");

        configIniFile.LoadIniFile();
        configIniFile.ShowIniFile();        

        Console.WriteLine("");
        Console.WriteLine("Digite enter para voltar ao menu anterior");
        Console.ReadLine();
            
        Console.Clear();
        CallStartMenu();
    }
    public void CallStartMenu()
    {
        Console.WriteLine("ATUALIZADOR DE SERVICOS PROTHEUS");
        Console.WriteLine("----------------------------------");
        Console.WriteLine("Duvidas: danilo.david@totvs.com.br");
        Console.WriteLine("----------------------------------");
        Console.WriteLine("");
        Console.WriteLine("1-Carregar as variaveis definidas no arquivo de configuracao");
        Console.WriteLine("2-Carregar a lista dos servicos a serem atualizados");
        Console.WriteLine("3-Iniciar o backup dos servicos");
        Console.WriteLine("4-Realizar a atualizacao dos servicos do protheus");
        Console.WriteLine("5-Parar todos os servicos");
        Console.WriteLine("6-Iniciar todos os servicos");
        Console.WriteLine("7-Atualizar os arquivos INI dos servicos");


        Console.WriteLine("");
        Console.WriteLine("");
        Console.WriteLine("INFORME QUAL OPCAO DESEJA SEGUIR:");

        string option = Console.ReadLine();
        //  return option;

        switch (option)
        {
            case "1":
                Console.Clear();
                LoadIniFile();
                break;

            case "2":
                Console.WriteLine("Segue a lista de servicos encontrados: ");
                Console.WriteLine("");
                windowsServicesManager.WriteListOfServices(configIniFile.GetSearchPatternList(), configIniFile.GetExcludePatternList());
                Console.WriteLine("Aperte enter para voltar o menu anterior: ");
                Console.WriteLine("");
                Console.ReadLine();
                Console.Clear();
                CallStartMenu();
                break;

            case "3":
                Console.WriteLine("Iniciando backup dos servicos: ");
                Console.WriteLine("");
                windowsServicesManager.BackupListWindowsServices();
                Console.WriteLine("Aperte enter para voltar o menu anterior: ");
                Console.WriteLine("");
                Console.ReadLine();
                Console.Clear();
                CallStartMenu();
                break;

            case "4":
                Console.WriteLine("Iniciando a atualizacao dos binarios: ");
                Console.WriteLine("");
                windowsServicesManager.UpdateDirectoryWindowsService(configIniFile.GetUpdateDir());
                Console.WriteLine("Aperte enter para voltar o menu anterior: ");
                Console.WriteLine("");
                Console.ReadLine();
                Console.Clear();
                CallStartMenu();
                break;

            case "5":
                Console.WriteLine("Parando os serviços do protheus: ");
                Console.WriteLine("");
                windowsServicesManager.StopListWindowsServices();
                Console.WriteLine("Aperte enter para voltar o menu anterior: ");
                Console.WriteLine("");
                Console.ReadLine();
                Console.Clear();
                CallStartMenu();
                break;
            

            case "6":
                Console.WriteLine("Iniciando os serviços do protheus: ");
                Console.WriteLine("");
                windowsServicesManager.StartListWindowsServices();
                Console.WriteLine("Aperte enter para voltar o menu anterior: ");
                Console.WriteLine("");
                Console.ReadLine();
                Console.Clear();
                CallStartMenu();
                break;

            case "7":
                Console.WriteLine("Iniciando a atualizacao dos Arquivos INI: ");
                Console.WriteLine("");
                windowsServicesManager.UpdateIniWindowsServices();
                Console.WriteLine("Aperte enter para voltar o menu anterior: ");
                Console.WriteLine("");
                Console.ReadLine();
                Console.Clear();
                CallStartMenu();
                break;

            case "8":
                Console.WriteLine("Iniciando o split do arquivo: ");
                Console.WriteLine("");
                windowsServicesManager.SplitLineOfFile();
                Console.WriteLine("Aperte enter para voltar o menu anterior: ");
                Console.WriteLine("");
                Console.ReadLine();
                Console.Clear();
                CallStartMenu();
                break;              


            default:
                Console.WriteLine("A opcao informada nao esta correta, aperte enter para voltar ao menu e tentar novamente!");
                Console.ReadLine();
                Console.Clear();
                CallStartMenu();
                break;
        }
    }
}

class ConfigIniFile
{    
    IniFile iniFile;

    private List<string> searchPatternList;
    private List<string> excludePatternList;
    
    private string searchPattern;
    private string excludePattern;
    private string backupDir;
    private string updateDir;

    public ConfigIniFile(string iniFileName)
    {
        //iniFileName = "AtualizadorServicosProtheus.ini"
        
        if (!File.Exists(iniFileName))
        {
            File.Create(iniFileName).Close();
        }

        iniFile = new IniFile(iniFileName);
    }

    public string GetVal(string key, string section)
    {       
        string stringVal = iniFile.Read(key, section);        
        return stringVal;
    }

    public List<string> ConvertStringToList(string text)
    {
        List<string> list = new List<string>(text.Split(';'));            
        return list;
    }

    public void LoadIniFile()
    {
        searchPattern = GetVal("searchPattern", "ProtheusServices");
        excludePattern = GetVal("excludePattern", "ProtheusServices");
        backupDir = GetVal("backupDir", "ProtheusServices");
        updateDir = GetVal("updateDir", "ProtheusServices");

        searchPatternList = ConvertStringToList(searchPattern);
        excludePatternList = ConvertStringToList(excludePattern);        
    }

    public void ShowIniFile()
    {
        Console.WriteLine("Padrao de busca: " + searchPattern);
        Console.WriteLine("Padrao de exclusao: " + excludePattern);
        Console.WriteLine("Diretorio de backup: " + backupDir);
        Console.WriteLine("Diretorio de origem do backup: " + updateDir);
    }

    public List<string> GetSearchPatternList()
    {
        return searchPatternList;
    }

    public List<string> GetExcludePatternList()
    {
        return excludePatternList;
    }

    public string GetBackupDir()
    {
        return backupDir;
    }

    public string GetUpdateDir()
    {
        return updateDir;
    }

    public void DeleteSection(string section)
    {
        iniFile.DeleteSection(section);
    }

    public void CreateKey(string key, string value, string section)
    {
        iniFile.Write(key, value, section);
    }
}

class WindowsServicesManager
{
    List<string[]> listOfServices;
    Folder folder;
    string resultFolder = "result";
    string logFolder = "log";

    public WindowsServicesManager()
    {
        listOfServices = new List<string[]>();
        folder = new Folder();
    }


    public void SplitLineOfFile()
    {
        using (StreamWriter writetext = new StreamWriter("C:\\Users\\vrdan\\Downloads\\Scripts PETZ\\Scripts PETZ\\query2_tratada.sql"))
        {
        //StreamReader readtext = new StreamReader("C:\\Users\\vrdan\\Downloads\\Scripts PETZ\\Scripts PETZ\\query2.sql");
        
            string line = File.ReadAllText("C:\\Users\\vrdan\\Downloads\\Scripts PETZ\\Scripts PETZ\\query2.sql");
            // readtext.ReadLine();

            string[] readtextArray = line.Split(',');

            foreach (string readtext in readtextArray)
            {

                //  Console.WriteLine(readtextList.Count);

                writetext.WriteLine("DELETE FROM SX5040 WHERE R_E_C_N_O_ = " + readtext.Trim() + ";" );
            }
            
        }
    }

    public void UpdateIniWindowsServices()
    {
        ConfigIniFile configIniFile;
        ConfigIniFile configIniFileServices;
        string iniFile = "AtualizadorServicosProtheus.ini";

        string[] services;
        folder.CreateIfNotExists(logFolder);
        string option = '0'.ToString();

        Console.WriteLine("");
        Console.WriteLine("Segue abaixo a lista dos INIs que serao atualizados");
        Console.WriteLine("");

        folder.CreateIfNotExists(resultFolder);
        string resultFile = resultFolder + "\\listOfINIBeUpdated.result";
        using (StreamWriter writetext = new StreamWriter(resultFile))
        {
            for (int i = 0; i < listOfServices.Count; i++)
            {
                services = listOfServices[i];
                Console.WriteLine(services[2] + "\\" + services[4]);
                writetext.WriteLine(services[2] + "\\" + services[4]);
            }
        }
        Console.WriteLine("");
        Console.WriteLine("Você também poderá consultar a lista através do arquivo: " + Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + resultFile);

        Console.WriteLine("");
        Console.WriteLine("Dado a lista dos INIs a serem atualizados acima, confirme o inicio da atualizacao: 1-Iniciar 2-Abortar");
        Console.WriteLine("");

        option = Console.ReadLine();
        if (option == '1'.ToString())
        {
            configIniFile = new ConfigIniFile(iniFile);

            string webAppSection = "WebApp";
            string webAgentSection = "WebAgent";
            string webAppKeyPort = "Port";            
            string webAppKeyEnable = "Enable";
            string webAppValueEnable = "1";
            string webAppKeyAgentJsonUpdate = "agentJsonUpdate";

            string webAppStartValuePort = configIniFile.GetVal("webAppStartValuePort", "webapp");
            string webAppValueAgentJsonUpdate = configIniFile.GetVal("webAppValueAgentJsonUpdate", "webapp"); 
            
            string originIniFile;
            string backupIniFile;
            string webAppValuePort = webAppStartValuePort;
            using (StreamWriter writetext = new StreamWriter(logFolder + "\\INIsAlterados" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".log"))
            {
                for (int i = 0; i < listOfServices.Count; i++)
                {
                    
                    services = listOfServices[i];
                    originIniFile = services[2] + "\\" + services[4];
                    backupIniFile = originIniFile + "_" + DateTime.Now.ToString("yyyyMMdd_HHmm");
                    Console.WriteLine("Realizando o backup do arquivo: " + originIniFile);
                    writetext.WriteLine("Realizando o backup do arquivo: " + originIniFile);
                    File.Copy(originIniFile, backupIniFile, true);

                    Console.WriteLine("Realizando a atualizacao do arquivo: " + originIniFile);
                    writetext.WriteLine("Realizando a atualizacao do arquivo: " + originIniFile);
                    configIniFileServices = new ConfigIniFile(originIniFile);
                    string val = configIniFileServices.GetVal(webAppKeyPort, webAppSection);
                                         
                    configIniFileServices.DeleteSection(webAppSection);
                    configIniFileServices.DeleteSection(webAgentSection);
                    if (val != "")
                    {
                        configIniFileServices.CreateKey(webAppKeyPort, val, webAppSection);
                    }
                    else
                    {
                        configIniFileServices.CreateKey(webAppKeyPort, webAppValuePort, webAppSection);
                        int result = Int32.Parse(webAppValuePort);
                        result += 1;
                        webAppValuePort = result.ToString();
                    }                    
                    configIniFileServices.CreateKey(webAppKeyEnable, webAppValueEnable, webAppSection);
                    configIniFileServices.CreateKey(webAppKeyAgentJsonUpdate, webAppValueAgentJsonUpdate, webAppSection);

                    Console.WriteLine("");
                    writetext.WriteLine("");
                    // limpar a sec webapp
                    // limpar a sec webagent
                    // incluir a sec webapp com o parametro
                    // [WebApp]
                    // Port = 4322
                    //Enable = 1
                    //agentJsonUpdate = D:\Totvs\Protheus\webagent\webagent.json                    
                }
            }
        }   
        else
        {
            Console.WriteLine("");
            Console.WriteLine("Atualizacao cancelada.");
            Console.WriteLine("");
        }
    }

    public void GenerateListWindowsServices(List<string> searchPattern, List<string> excludePattern)
    {
        listOfServices.Clear();

        ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Service WHERE StartMode <> 'Disabled'");
        ManagementObjectCollection collection = searcher.Get();

        int countExclude=0;

        // Para cada servico da lista de servicos do windows
        foreach (ManagementObject service in collection)
        {
            // Para cada valor indicado no parametro searchPattern para encontrar os servicos
            foreach (string include in searchPattern)
            {
                if (include != "")
                { 
                    // Verifica se o nome do servico do windows contem algum dos valores indicados no parametro serachPattern
                    if (service["DisplayName"].ToString().ToUpper().Contains(include.ToUpper()))
                    {
                        countExclude = 0;
                        // Para cada item da lista de valores a que nao devem ser encontrados nos servicos
                        foreach (string exclude in excludePattern)
                        {
                            if (exclude != "")
                            {
                                // Se o algum dos valores que nao devem constar na lista de servicos for identificado o contador sera colocado em 1
                                
                                if (service["DisplayName"].ToString().ToUpper().Contains(exclude.ToUpper()))
                                {
                                    countExclude = 1;
                                }
                            }
                        }

                        // Se o contador ainda for 0 indicando que nao deve ser excluido o servico e adicionado a lista final
                        if (countExclude == 0)
                        {
                            string serviceDirectory = Path.GetDirectoryName(service["PathName"].ToString());
                            string serviceExec = Path.GetFileName(service["PathName"].ToString());
                            string serviceIni = GetIniServiceName(serviceExec);
                            listOfServices.Add(new string[] { service["DisplayName"].ToString(), service["Name"].ToString(), serviceDirectory, serviceExec, serviceIni});
                        }
                    }
                }
            }
        }               
    }

    public void WriteListOfServices(List<string> searchPattern, List<string> excludePattern)
    {

        GenerateListWindowsServices(searchPattern, excludePattern);

        folder.CreateIfNotExists(resultFolder);

        string[] services;
        using (StreamWriter writetext = new StreamWriter(resultFolder+"/listOfServices.result"))
        {
            for (int i = 0; i < listOfServices.Count; i++)
            {
                services = listOfServices[i];
                Console.WriteLine("Displayname: " + services[0]);
                writetext.WriteLine(services[0]);
                // Console.WriteLine("Name: " + servicos[1]);
                // Console.WriteLine("PathNameDirectory: " + servicos[2]);
                // Console.WriteLine("PathNameExec: " + servicos[3]);
                // Console.WriteLine("");
                // Console.WriteLine("");
            }
        }

        using (StreamWriter writetext = new StreamWriter(resultFolder + "\\listOfDirServices.result"))
        {
            for (int i = 0; i < listOfServices.Count; i++)
            {
                services = listOfServices[i];
                writetext.WriteLine(services[2]);
            }
        }

        using (StreamWriter writetext = new StreamWriter(resultFolder + "\\listOfIniServices.result"))
        {
            for (int i = 0; i < listOfServices.Count; i++)
            {
                services = listOfServices[i];
                writetext.WriteLine(services[2]+"\\"+services[3]);
                writetext.WriteLine(services[2] + "\\" + services[4]);
                writetext.WriteLine("");
            }
        }

        Console.WriteLine("");
        Console.WriteLine("Total de servicos encontrados: " + listOfServices.Count);
        Console.WriteLine("");
    }
    
    public void UpdateServiceDirectory(string dirOrigin, string dirDestiny, string execName)
    {
        folder.CopyFolder(dirOrigin, dirDestiny);
        folder.RenameFile(dirDestiny, "appserver.exe", execName);
    }

    public void UpdateDirectoryWindowsService(string dirOrigin)
    {
        if (!Directory.Exists(dirOrigin))
        {

            Console.WriteLine("");
            Console.WriteLine("O diretorio de atualizacao nao foi encontrado. Por favor, crie a pasta atualizacao e insira dentro dela os componentes a serem atualizados!!!");
            Console.WriteLine("");
        }
        else
        {
            string[] services;
            string option = '0'.ToString();

            Console.WriteLine("");
            Console.WriteLine("Segue abaixo a lista dos diretorios que serao atualizados");
            Console.WriteLine("");

            folder.CreateIfNotExists(resultFolder);
            string resultFile = resultFolder + "\\listOfDirBeUpdated.result";
            using (StreamWriter writetext = new StreamWriter(resultFile))
            {
                for (int i = 0; i < listOfServices.Count; i++)
                {
                    services = listOfServices[i];
                    Console.WriteLine(services[2]);
                    writetext.WriteLine(services[2]);
                }
            }
            Console.WriteLine("Você também poderá consultar a lista através do arquivo: " + Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + resultFile);

            Console.WriteLine("");
            Console.WriteLine("Dado a lista dos diretorios a serem atualizados acima, confirme o inicio da atualizacao: 1-Iniciar 2-Abortar");
            Console.WriteLine("");

            option = Console.ReadLine();
            if (option == '1'.ToString())
            {
                folder.CreateIfNotExists(logFolder);
                using (StreamWriter writetext = new StreamWriter(logFolder + "\\AtualizacaoServicos_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".log"))
                {
                    for (int i = 0; i < listOfServices.Count; i++)
                    {
                        services = listOfServices[i];
                        Console.WriteLine("Realizando a limpeza do diretorio: " + services[2]);
                        CleanDirectory(services[2], GetIniServiceName(services[3]));
                        writetext.WriteLine("Realizando a limpeza do diretorio: " + services[2]);

                        Console.WriteLine("Realizando a copia dos novos arquivos do diretorio " + dirOrigin + " para o diretorio " + services[2]);
                        UpdateServiceDirectory(dirOrigin, services[2], services[3]);
                        writetext.WriteLine("Realizando a copia dos novos arquivos do diretorio " + dirOrigin + " para o diretorio " + services[2]);

                        Console.WriteLine("");
                        writetext.WriteLine("");
                    }
                }
            }
            else
            {
                Console.WriteLine("");
                Console.WriteLine("Atualizacao cancelada.");
                Console.WriteLine("");
            }
        }
    }

    public void CleanDirectory(string serviceDir, string exceptionFile)
    {
        foreach (string file in Directory.GetFiles(serviceDir))
        {
            if (Path.GetFileName(file) != exceptionFile)
            {
                File.Delete(file);
            }
        }

        foreach (string subDirectory in Directory.GetDirectories(serviceDir))
        {
            Directory.Delete(subDirectory, true);
        }
    }

    public string GetIniServiceName(string execServiceName)
    {
        string iniServiceName = Path.GetFileNameWithoutExtension(execServiceName) + ".ini";        
        return iniServiceName;
    }

    public void BackupListWindowsServices()
    {
        //folder = new Folder();
        string[] servicos;
        string unit;
        string unitWithoutColon;
        string pathWithoutUnit;
        string backupDir = "backup";
        string dirDestiny;

        if (Directory.Exists(backupDir))
        {
            // Criar a pasta de destino se não existir
            string backupDirOld = backupDir + DateTime.Now.ToString("yyyyMMdd_HHmm");
            Directory.Move(backupDir, backupDirOld);
        }

        for (int i = 0; i < listOfServices.Count; i++)
        {
            servicos = listOfServices[i];
            unit = Path.GetPathRoot(servicos[2]);
            unitWithoutColon = unit.Replace(":", "");
            pathWithoutUnit = servicos[2].Substring(unit.Length);
            dirDestiny = backupDir + '\\' + unitWithoutColon + pathWithoutUnit;

            folder.CopyFolder(servicos[2], dirDestiny.ToString());
        }        
    }

    public void StopListWindowsServices()
    {
        string[] services;
        for (int i = 0; i < listOfServices.Count; i++)
        {
            services = listOfServices[i];
            StopService(services[1]);
        }
    }

    public void StartListWindowsServices()
    {
        string[] services;
        for (int i = 0; i < listOfServices.Count; i++)
        {
            services = listOfServices[i];
            StartService(services[1]);
        }
    }

    static void StopService(string serviceName)
    {

        ServiceController service = new ServiceController(serviceName);
        if (service.Status == ServiceControllerStatus.Running)
        {
            service.Stop();
            service.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(10));
        }
    }

    static void StartService(string serviceName)
    {
        ServiceController service = new ServiceController(serviceName);
        if (service.Status == ServiceControllerStatus.Stopped)
        {
            service.Start();
        }
    }

    static void Main(string[] args)
    {
        Menu menu = new Menu();
        menu.CallStartMenu();
    }
}



