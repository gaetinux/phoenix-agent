using System;
using System.Diagnostics;
using System.IO;
using System.ServiceProcess;
using System.Threading;
using System.Threading.Tasks;
using System.Data.SQLite;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.Text.Json;
using Newtonsoft.Json;

namespace PhoenixAgent
{
    public class Session
    {
        public string User { get; set; }
        public string Protocol { get; set; }
        public string Id { get; set; }
        public string State { get; set; }
        public string InactiveFrom { get; set; }
        public string Logon { get; set; }
    }

    public class Program
    {
        public static async Task Main(string[] args)
        {
            if (args.Length > 0 && args[0] == "--install")
            {
                InstallService();
                return;
            }

            if (args.Length > 0 && args[0] == "--uninstall")
            {
                UninstallService();
                return;
            }

            var host = Host.CreateDefaultBuilder(args)
                .UseWindowsService()
                .ConfigureServices(services =>
                {
                    services.AddHostedService<Worker>();
                })
                .Build();

            await host.RunAsync();
        }

        private static void InstallService()
        {
            string exePath = Process.GetCurrentProcess().MainModule.FileName;
            string serviceName = "phoenix-agent";

            using var serviceController = new ServiceController(serviceName);
            try
            {
                if (serviceController.Status == ServiceControllerStatus.Stopped)
                {
                    Console.WriteLine("Le service existe déjà.");
                    return;
                }
            }
            catch
            {
                // Si le service n'existe pas, continue l'installation
            }

            Process.Start("sc", $"create {serviceName} binPath= \"{exePath}\" start= auto");
            Console.WriteLine("Service installé avec succès !");
        }

        private static void UninstallService()
        {
            string serviceName = "phoenix-agent";

            Process.Start("sc", $"delete {serviceName}");
            Console.WriteLine("Service désinstallé avec succès !");
        }
    }

    public class Worker : BackgroundService
    {
        private readonly string _databasePath = @"C:\ProgramData\phoenix-agent\sessions.db";

        public Worker()
        {
            InitializeDatabase();
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    string queryResult = ExecuteQueryUser();
                    Log(queryResult);
                    UpdateDatabase(queryResult);



                    Log("Résultat enregistré avec succès.");
                    await Task.Delay(TimeSpan.FromSeconds(20), stoppingToken);
                }
                catch (Exception ex)
                {
                    Log($"Erreur : {ex.Message}");
                }
            }
        }

        private void InitializeDatabase()
        {
            if (!File.Exists(_databasePath))
            {
                using var connection = new SQLiteConnection($"Data Source={_databasePath}");
                connection.Open();

                using var command = connection.CreateCommand();
                command.CommandText = @"
                    CREATE TABLE IF NOT EXISTS Sessions (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        User TEXT NOT NULL,
                        Protocol TEXT NOT NULL,
                        State TEXT NOT NULL,
                        InactiveFrom TEXT NOT NULL,
                        Logon TEXT NOT NULL
                    );";
                command.ExecuteNonQuery();

                Log("Base de données SQLite initialisée.");
            }
        }

        private string ExecuteQueryUser()
        {
            // Exemple d'un script PowerShell multiligne
            string script = @"
                $ProgressPreference = 'SilentlyContinue'

                # Initialisation de la liste de sessions
                $list_sessions = New-Object System.Collections.Generic.List[PSObject]

                # Récupérer la sortie de 'query user' une seule fois
                $queryOutput = query user

                # Parcours de chaque ligne de la sortie de 'query user'
                foreach($line in $queryOutput -split '\n'){
                    # Ignorer les lignes contenant des informations inutiles
                    if($line.Contains('UTILISATEUR') -or $line.Contains('Aucun') -or [string]::IsNullOrEmpty($line)) {
                        continue
                    }

                    # Séparation des colonnes
                    $parsed_line = $line -split '\s+'

                    # Si la première colonne est vide, ignorer cette ligne
                    if(![string]::IsNullOrEmpty($parsed_line[0])) { 
                        continue
                    }

                    # Création de l'objet session
                    $session = [pscustomobject]@{
                        User = $parsed_line[1]
                        Protocol = $parsed_line[2]
                        Id = $parsed_line[3]
                        State = $parsed_line[4]
                        InactiveFrom = $parsed_line[5]
                        Logon = $parsed_line[6]
                    }

                    # Ajouter la session à la liste
                    $list_sessions.Add($session)
                }

                # Afficher les sessions collectées
                $list_sessions | ConvertTo-Json -Compress
            ";

            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-ExecutionPolicy Bypass -Command \"{script}\"",  // Passe le script complet
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            process.Start();
            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();
            process.WaitForExit();

            if (process.ExitCode != 0)
            {
                throw new Exception($"Erreur lors de l'exécution du script PowerShell : {error}");
            }

            return output;  // Retourne la sortie du script
        }


        private void UpdateDatabase(string jsonResult)
        {
            // Désérialiser le JSON en une liste d'objets Session
            var sessions = JsonConvert.DeserializeObject<List<Session>>(jsonResult);

            // Ajouter les nouvelles sessions à la base de données
            using var connection = new SQLiteConnection($"Data Source={_databasePath}");
            connection.Open();

            using var deleteCommand = connection.CreateCommand();
            deleteCommand.CommandText = "DELETE FROM Sessions;";
            deleteCommand.ExecuteNonQuery();

            foreach (var session in sessions)
            {
                using var command = connection.CreateCommand();
                command.CommandText = @"
                    INSERT INTO Sessions (Id, User, Protocol, State, InactiveFrom, Logon)
                    VALUES (@Id, @User, @Protocol, @State, @InactiveFrom, @Logon);";

                command.Parameters.AddWithValue("@Id", session.Id);
                command.Parameters.AddWithValue("@User", session.User);
                command.Parameters.AddWithValue("@Protocol", session.Protocol);
                command.Parameters.AddWithValue("@State", session.State);
                command.Parameters.AddWithValue("@InactiveFrom", session.InactiveFrom);
                command.Parameters.AddWithValue("@Logon", session.Logon);

                command.ExecuteNonQuery();
            }
        }

        private void Log(string message)
        {
            string logFile = @"C:\ProgramData\phoenix-agent\service.log";
            Directory.CreateDirectory(Path.GetDirectoryName(logFile));
            File.AppendAllText(logFile, $"{DateTime.Now}: {message}{Environment.NewLine}");
        }
    }
}
