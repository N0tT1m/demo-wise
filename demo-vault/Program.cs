using DemoFile;
using System.Drawing;
using System.Drawing.Imaging;
using ClosedXML.Excel;
using System.Diagnostics;
using DemoFile.Game.Cs;
using DemoFile.Sdk;
using Spectre.Console;
using Color = System.Drawing.Color;
using System.Data;
using System.Data.SqlClient;
using Microsoft.Data.SqlClient;

internal class MapConfig
{
    public string Name { get; set; }
    public float PosX { get; set; }
    public float PosY { get; set; }
    public float Scale { get; set; }
    public bool HasLowerLevel { get; set; }
    public float DefaultAltitudeMax { get; set; }
    public float DefaultAltitudeMin { get; set; }
    public float LowerAltitudeMax { get; set; }
    public float LowerAltitudeMin { get; set; }
}

// The main class that holds round timing information
public class RoundInfo
{
    public int StartTick { get; set; }
    public int FreezeEndTick { get; set; }
    public int EndTick { get; set; }
}

// Class to represent a complete round with all its metadata
public class ValidRound
{
    public int RoundNumber { get; set; }
    public string MVP { get; set; }
    public int? MVPReason { get; set; }
    public bool IsMatchPoint { get; set; }
    public bool IsFinal { get; set; }
    public bool IsLastRoundHalf { get; set; }
}

internal class DemoInfo
{
    public string DemoName { get; set; } // Original demo file name
    public string MatchId { get; set; }  // Match identifier for grouping related demos
    public string MapName { get; set; }  // The map being played
    public DateTime RecordedAt { get; set; } // When the demo was recorded
}

internal class Program
{ 
    // Add this near the top of the Program class with other private fields
    private static readonly string _connectionString = BuildConnectionString();

    // Add this method to the Program class to construct the connection string
    private static string BuildConnectionString()
    {
        // You can modify these values based on your SQL Server configuration
        var builder = new SqlConnectionStringBuilder
        {
            DataSource = "192.168.1.78,1433", // Server name and port
            InitialCatalog = "DemoStats", // Database name
            UserID = "sa",         // SQL Server username
            Password = "B@bycakes15!",    // SQL Server password
            
            // Additional recommended settings for better security and performance
            TrustServerCertificate = true,    // Required for newer SQL Server versions
            Encrypt = true,                   // Enable encryption
            ConnectTimeout = 30,              // Connection timeout in seconds
            ApplicationName = "demo-vault"    // Helps identify your application in SQL Server logs
        };

        return builder.ConnectionString;
    }
    
    private static readonly Dictionary<string, MapConfig> MapConfigs = new()
    {
        {
            "de_ancient", new MapConfig
            {
                Name = "ancient",
                PosX = -2953,
                PosY = 2164,
                Scale = 5
            }
        },
        {
            "de_anubis", new MapConfig
            {
                Name = "anubis",
                PosX = -2796,
                PosY = 3328,
                Scale = 5.22f
            }
        },
        {
           "de_dust2", new MapConfig
           {
               Name = "dust2",
               PosX = -2476,
               PosY = 3239,
               Scale = 4.4f
           }
        },
        {
            "de_inferno", new MapConfig
            {
                Name = "inferno",
                PosX = -2087,
                PosY = 3870,
                Scale = 4.9f
            }
        },
        {
            "de_mirage", new MapConfig
            {
                Name = "mirage",
                PosX = -3230,
                PosY = 1713,
                Scale = 5.00f,
            }
        },
        {
            "de_nuke", new MapConfig
            {
                Name = "nuke",
                PosX = -3453,
                PosY = 2887,
                Scale = 7f,
                HasLowerLevel = true,
                DefaultAltitudeMax = 10000,
                DefaultAltitudeMin = -495,
                LowerAltitudeMax = -495,
                LowerAltitudeMin = -10000
            }
        },
        {
            "de_overpass", new MapConfig
            {
                Name = "overpass",
                PosX = -4831,
                PosY = 1781,
                Scale = 5.20f,
            }
        },
        {
            "de_train", new MapConfig
            {
                Name = "train",
                PosX = -2308,
                PosY = 2078,
                Scale = 4.082077f,
                HasLowerLevel = true,
                DefaultAltitudeMax = 20000,
                DefaultAltitudeMin = -50,
                LowerAltitudeMax = -50,
                LowerAltitudeMin = -5000
            }
        },
        {
            "de_vertigo", new MapConfig
            {
                Name = "vertigo",
                PosX = -3168,
                PosY = 1762,
                Scale = 4.0f,
                HasLowerLevel = true,
                DefaultAltitudeMax = 20000,
                DefaultAltitudeMin = 11700,
                LowerAltitudeMax = 11700,
                LowerAltitudeMin = -10000
            }
        },
        // Add other maps as needed
    };

    // Dictionary to store grenades by type
    private static Dictionary<string, List<(float X, float Y, float Z, string Player, string DemoId)>> grenadePositions = new()
    {
        { "Smoke", new List<(float, float, float, string, string)>() },
        { "Molotov", new List<(float, float, float, string, string)>() },
        { "HE", new List<(float, float, float, string, string)>() },
        { "Flashbang", new List<(float, float, float, string, string)>() }
    };
    
    // Add these new data structures at the class level, alongside the existing grenadePositions dictionary
    private static List<(int roundNum, uint Value, string, string, string Weapon)> grenadeEvents = new();
    private static List<(int roundNum, uint Value, string PlayerName, List<(string GrenadeType, int Count)> inventory)> grenadeInventory = new();
    
    // Update the data structure to store round information
    private static Dictionary<int, RoundInfo> roundTicks = new();

    // Create a class to store round information
    internal class RoundInfo
    {
        public int StartTick { get; set; }
        public int FreezeEndTick { get; set; }
        public int EndTick { get; set; }
        public string MVP { get; set; } = string.Empty;
        public int MvpReason { get; set; }
        public bool IsMatchPoint { get; set; }
        public bool IsFinal { get; set; }
        public bool IsLastRoundHalf { get; set; }
    }
    
    // Add this class to store player statistics
    internal class PlayerStats
    {
        public int Round { get; set; }
        public string PlayerName { get; set; }
        public string SteamID { get; set; }
        public string Team { get; set; }
        public int Ping { get; set; }
        public int Money { get; set; }
        public int Kills { get; set; }
        public int Deaths { get; set; }
        public int Assists { get; set; }
        public double HeadshotPercentage { get; set; }
        public int Damage { get; set; }
    }

    // Add this list to store player statistics for each round
    private static List<PlayerStats> playerRoundStats = new();
    
    // Add this class with the other event tracking classes
    internal class DeathEvent
    {
        public int RoundNumber { get; set; }
        public uint Tick { get; set; }
        public string AttackerName { get; set; }
        public string AttackerTeam { get; set; }
        public string VictimName { get; set; }
        public string VictimTeam { get; set; }
        public string Weapon { get; set; }
        public bool Headshot { get; set; }
    }

    // Add this to your class-level fields
    private static List<DeathEvent> deathEvents = new();

    private static MapConfig currentMap;
    private static int currentRoundNumber = 0;
    // Add this field at the class level
    private static DemoInfo currentDemoInfo;
    
    private static void CapturePlayerStats(CsDemoParser demo, int roundNum, string demoName)
    {
        Console.WriteLine($"\nCapturing stats for round {roundNum}");
        
        foreach (var team in new[] { demo.TeamCounterTerrorist, demo.TeamTerrorist })
        {
            if (team?.CSPlayerControllers == null) 
            {
                Console.WriteLine("  Team or controllers is null");
                continue;
            }

            foreach (var player in team.CSPlayerControllers)
            {
                if (player == null) 
                {
                    Console.WriteLine("  Player is null");
                    continue;
                }

                var matchStats = player.ActionTrackingServices?.MatchStats ?? new CSMatchStats();
                var teamName = player.CSTeamNum switch
                {
                    CSTeamNumber.Terrorist => "T",
                    CSTeamNumber.CounterTerrorist => "CT",
                    _ => "SPEC"
                };

                // Calculate headshot percentage safely
                double headshotPercentage = 0;
                if (matchStats.Kills > 0 && matchStats.HeadShotKills >= 0)
                {
                    headshotPercentage = (matchStats.HeadShotKills * 100.0) / matchStats.Kills;
                }
                
                playerRoundStats.Add(new PlayerStats
                {
                    Round = roundNum,
                    PlayerName = player.PlayerName ?? "Unknown",
                    SteamID = player.SteamID.ToString() ?? "Unknown",
                    Team = teamName,
                    Ping = (int)(player.Ping),
                    Money = player.InGameMoneyServices?.Account ?? 0,
                    Kills = matchStats.Kills,
                    Deaths = matchStats.Deaths,
                    Assists = matchStats.Assists,
                    HeadshotPercentage = headshotPercentage,
                    Damage = matchStats.Damage
                });

                Console.WriteLine($"  Added stats for {player.PlayerName}: Round {roundNum}, Team {teamName}, K/D/A: {matchStats.Kills}/{matchStats.Deaths}/{matchStats.Assists}");
            }
        }
    }
    
    private static string TeamNumberToString(CSTeamNumber? csTeamNumber) => csTeamNumber switch
    {
        CSTeamNumber.Terrorist => "T",
        CSTeamNumber.CounterTerrorist => "CT",
        _ => "Spec",
    };

    // Add this method to track whether a round is a warmup round
    private static bool IsWarmupRound(CsDemoParser demo)
    {
        return demo.GameRules?.WarmupPeriod ?? false;
    }
    
    // Update CleanupRoundNumbers method to handle gaps
    private static void CleanupRoundNumbers()
    {
        Console.WriteLine("\nStarting round number cleanup...");
        
        // First, remove warmup rounds (round 0)
        playerRoundStats.RemoveAll(stat => stat.Round <= 0);
        Console.WriteLine("Removed warmup round stats");

        // Get all unique round numbers in order
        var roundNumbers = playerRoundStats
            .Select(stat => stat.Round)
            .Distinct()
            .OrderBy(r => r)
            .ToList();

        Console.WriteLine($"Original round numbers: {string.Join(", ", roundNumbers)}");

        // Create sequential mapping starting from 1
        var sequentialMapping = new Dictionary<int, int>();
        for (int i = 0; i < roundNumbers.Count; i++)
        {
            sequentialMapping[roundNumbers[i]] = i + 1;
            Console.WriteLine($"Mapping round {roundNumbers[i]} to {i + 1}");
        }

        // Update round numbers in all data structures
        foreach (var stat in playerRoundStats)
        {
            var oldRound = stat.Round;
            stat.Round = sequentialMapping[oldRound];
        }

        // Update round numbers in roundTicks
        var updatedRoundTicks = new Dictionary<int, RoundInfo>();
        foreach (var (round, info) in roundTicks)
        {
            if (sequentialMapping.ContainsKey(round))
            {
                updatedRoundTicks[sequentialMapping[round]] = info;
            }
        }
        roundTicks.Clear();
        foreach (var (round, info) in updatedRoundTicks)
        {
            roundTicks[round] = info;
        }

        // Update round numbers in grenade events using a new list
        grenadeEvents = grenadeEvents.Select(evt => 
            sequentialMapping.ContainsKey(evt.roundNum) 
                ? (sequentialMapping[evt.roundNum], evt.Value, evt.Item3, evt.Item4, evt.Weapon)
                : evt
        ).ToList();

        // Update round numbers in grenade inventory using a new list
        grenadeInventory = grenadeInventory.Select(inv =>
            sequentialMapping.ContainsKey(inv.roundNum)
                ? (sequentialMapping[inv.roundNum], inv.Value, inv.PlayerName, inv.inventory)
                : inv
        ).ToList();

        Console.WriteLine($"Normalized round numbers: {string.Join(", ", sequentialMapping.Values.OrderBy(x => x))}");
    }

    public static async Task Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: program <demo_path> <map_name>");
            return;
        }

        var demoPath = args[0];
        var mapName = args[1].ToLower();

        if (!MapConfigs.ContainsKey(mapName))
        {
            Console.WriteLine($"Unknown map: {mapName}");
            return;
        }

        currentMap = MapConfigs[mapName];
        Console.WriteLine($"\nProcessing demo for map: {mapName}");
        if (currentMap.HasLowerLevel)
        {
            Console.WriteLine($"Map has lower level. Split at Z: {currentMap.LowerAltitudeMax}");
            Console.WriteLine(
                $"Upper level Z range: {currentMap.DefaultAltitudeMin} to {currentMap.DefaultAltitudeMax}");
            Console.WriteLine($"Lower level Z range: {currentMap.LowerAltitudeMin} to {currentMap.LowerAltitudeMax}");
        }

        var demoName = Path.GetFileName(demoPath);

        var demo = new CsDemoParser();
        var cts = new CancellationTokenSource();

        // Initialize demo info early
        currentDemoInfo = new DemoInfo
        {
            DemoName = Path.GetFileName(demoPath),
            MatchId = Convert.ToBase64String(
                System.Security.Cryptography.SHA256.HashData(
                    System.Text.Encoding.UTF8.GetBytes(demoPath)
                )).Substring(0, 20),
            MapName = mapName,
            RecordedAt = File.GetLastWriteTime(demoPath)
        };

        // Update the event handler with corrected round numbering
        demo.Source1GameEvents.RoundStart += e =>
        {
            if (IsWarmupRound(demo)) return; // Skip warmup rounds

            // Initialize first round as 1 instead of incrementing
            if (currentRoundNumber == 0)
            {
                currentRoundNumber = 1;
            }
            else
            {
                currentRoundNumber += 1;
            }

            Console.WriteLine($"\n\n>>> Round start [{currentRoundNumber}] <<<");

            // Initialize round info
            roundTicks[currentRoundNumber] = new RoundInfo
            {
                StartTick = (int)demo.CurrentGameTick.Value,
            };

            // Set end tick for previous round if it exists
            if (currentRoundNumber > 1 && roundTicks.ContainsKey(currentRoundNumber - 1))
            {
                var prevRound = roundTicks[currentRoundNumber - 1];
                if (prevRound.EndTick == 0)
                {
                    prevRound.EndTick = (int)demo.CurrentGameTick.Value;
                    roundTicks[currentRoundNumber - 1] = prevRound;
                }
            }
        };

        demo.Source1GameEvents.SmokegrenadeDetonate += e =>
        {
            // if (e.Player?.PlayerName != "--uncle-") return;
            grenadePositions["Smoke"].Add((e.X, e.Y, e.Z, e.Player?.PlayerName ?? "Unknown", demoPath));
            // Console.WriteLine($"{e.Player?.PlayerName} [Smoke] - [{e.X} {e.Y} {e.Z}]; Level: {(e.Z <= currentMap.LowerAltitudeMax ? "Lower" : "Upper")}");
        };

        demo.Source1GameEvents.InfernoStartburn += e =>
        {
            // if (e.Player?.PlayerName != "--uncle-") return;
            grenadePositions["Molotov"].Add((e.X, e.Y, e.Z, "Not Available", demoPath));
            // Console.WriteLine($"[Molotov] - [{e.X} {e.Y} {e.Z}] Level: {(e.Z <= currentMap.LowerAltitudeMax ? "Lower" : "Upper")}");
        };

        demo.Source1GameEvents.HegrenadeDetonate += e =>
        {
            // if (e.Player?.PlayerName != "--uncle-") return;
            grenadePositions["HE"].Add((e.X, e.Y, e.Z, e.Player?.PlayerName ?? "Unknown", demoPath));
            // Console.WriteLine($"{e.Player?.PlayerName} [HE] - [{e.X} {e.Y} {e.Z}] Level: {(e.Z <= currentMap.LowerAltitudeMax ? "Lower" : "Upper")}");
        };

        demo.Source1GameEvents.FlashbangDetonate += e =>
        {
            // if (e.Player?.PlayerName != "--uncle-") return;
            grenadePositions["Flashbang"].Add((e.X, e.Y, e.Z, e.Player?.PlayerName ?? "Unknown", demoPath));
            // Console.WriteLine($"{e.Player?.PlayerName} [Flashbang] - [{e.X} {e.Y} {e.Z}] Level: {(e.Z <= currentMap.LowerAltitudeMax ? "Lower" : "Upper")}");
        };

        // Update death event handler
        demo.Source1GameEvents.PlayerDeath += e =>
        {
            if (IsWarmupRound(demo)) return;

            var deathEvent = new DeathEvent
            {
                RoundNumber = currentRoundNumber,
                Tick = demo.CurrentGameTick.Value,
                AttackerName = e.Attacker?.PlayerName ?? "Unknown",
                AttackerTeam = TeamNumberToString(e.Attacker?.CSTeamNum),
                VictimName = e.Player?.PlayerName ?? "Unknown",
                VictimTeam = TeamNumberToString(e.Player?.CSTeamNum),
                Weapon = e.Weapon,
                Headshot = e.Headshot
            };

            deathEvents.Add(deathEvent);

            // AnsiConsole.Markup($"[{deathEvent.AttackerTeam}]{deathEvent.AttackerName}[/]");
            // AnsiConsole.Markup(" <");
            // AnsiConsole.Markup(deathEvent.Weapon);
            // if (deathEvent.Headshot)
            //     AnsiConsole.Markup(" HS");
            // AnsiConsole.Markup("> ");
            // AnsiConsole.MarkupLine($"[{deathEvent.VictimTeam}]{deathEvent.VictimName}[/]");
        };

        demo.Source1GameEvents.RoundFreezeEnd += e =>
        {
            Console.WriteLine("\n  > Round freeze end");
            if (roundTicks.ContainsKey(currentRoundNumber))
            {
                var record = roundTicks[currentRoundNumber];
                record.FreezeEndTick = (int)demo.CurrentGameTick.Value;
            }

            DumpGrenadeInventory();
        };

        // Add this with the other event handlers in the Main method
        demo.Source1GameEvents.RoundEnd += e =>
        {
            Console.WriteLine("\n  > Round end");
            if (roundTicks.ContainsKey(currentRoundNumber))
            {
                var record = roundTicks[currentRoundNumber];
                record.EndTick = (int)demo.CurrentGameTick.Value;
            }

            // Make sure we capture stats at round end
            CapturePlayerStats(demo, currentRoundNumber, demoPath);
            DumpGrenadeInventory();
        };

        demo.Source1GameEvents.RoundMvp += e =>
        {
            Console.WriteLine($"\n  > Round MVP: {e.Player?.PlayerName ?? "Unknown"} (Reason: {e.Reason})");
            if (roundTicks.ContainsKey(currentRoundNumber))
            {
                roundTicks[currentRoundNumber].MVP = e.Player?.PlayerName ?? "Unknown";
                roundTicks[currentRoundNumber].MvpReason = e.Reason;
                Console.WriteLine($"  Stored MVP for round {currentRoundNumber}: {roundTicks[currentRoundNumber].MVP}");
            }
        };

        demo.Source1GameEvents.RoundAnnounceMatchPoint += e =>
        {
            Console.WriteLine($"\n  > Match Point announced for round {currentRoundNumber}");
            if (roundTicks.ContainsKey(currentRoundNumber))
            {
                roundTicks[currentRoundNumber].IsMatchPoint = true;
                Console.WriteLine($"  Set match point flag for round {currentRoundNumber}");
            }
        };

        demo.Source1GameEvents.RoundAnnounceFinal += e =>
        {
            Console.WriteLine($"\n  > Final Round announced for round {currentRoundNumber}");
            if (roundTicks.ContainsKey(currentRoundNumber))
            {
                roundTicks[currentRoundNumber].IsFinal = true;
                Console.WriteLine($"  Set final round flag for round {currentRoundNumber}");
            }
        };

        demo.Source1GameEvents.RoundAnnounceLastRoundHalf += e =>
        {
            Console.WriteLine($"\n  > Last Round in Half announced for round {currentRoundNumber}");
            if (roundTicks.ContainsKey(currentRoundNumber))
            {
                roundTicks[currentRoundNumber].IsLastRoundHalf = true;
                Console.WriteLine($"  Set last round in half flag for round {currentRoundNumber}");
            }
        };

        demo.EntityEvents.CCSPlayerPawn.AddCollectionChangeCallback(pawn => pawn.Grenades,
            (pawn, oldGrenades, newGrenades) =>
            {
                Console.Write($"  [Tick {demo.CurrentGameTick.Value}] ");
                MarkupPlayerName(pawn.Controller);
                var playerName = pawn.Controller?.PlayerName ?? "Unknown";

                // Store the grenade changes
                grenadeEvents.Add((
                    currentRoundNumber,
                    demo.CurrentGameTick.Value,
                    playerName,
                    "InventoryChange",
                    string.Join(", ", newGrenades.Select(x => x.ServerClass.Name))
                ));

                AnsiConsole.MarkupLine(
                    $" grenades changed [grey]{string.Join(", ", oldGrenades.Select(x => x.ServerClass.Name))}[/] => [bold]{string.Join(", ", newGrenades.Select(x => x.ServerClass.Name))}[/]");
            });

        // Update grenade event handler - consider only storing events during actual gameplay rounds
        demo.Source1GameEvents.WeaponFire += e =>
        {
            if (!e.Weapon.Contains("nade") && !e.Weapon.Contains("molotov"))
                return;

            if (IsWarmupRound(demo)) return;

            Console.Write($"  [Tick {demo.CurrentGameTick.Value}] ");
            MarkupPlayerName(e.Player);

            grenadeEvents.Add((
                currentRoundNumber,
                demo.CurrentGameTick.Value,
                e.Player?.PlayerName ?? "Unknown",
                "Throw",
                e.Weapon
            ));

            AnsiConsole.MarkupLine($" [bold]threw a {e.Weapon}[/]");
        };

        // Update inventory tracking
        void DumpGrenadeInventory()
        {
            if (IsWarmupRound(demo)) return;

            foreach (var player in demo.Players)
            {
                var inventory = new List<(string GrenadeType, int Count)>();

                if (player.PlayerPawn is not { } pawn)
                {
                    if (player.PlayerName == "SourceTV") continue;
                    inventory.Add(("NO_PAWN", 0));
                }
                else if (!pawn.IsAlive)
                {
                    inventory.Add(("DEAD", 0));
                }
                else
                {
                    var grenades = pawn.Grenades;
                    if (!grenades.Any())
                    {
                        inventory.Add(("NO_GRENADES", 0));
                    }
                    else
                    {
                        foreach (var grenade in grenades)
                        {
                            inventory.Add((grenade.ServerClass.Name, grenade.GrenadeCount));
                        }
                    }
                }

                grenadeInventory.Add((
                    currentRoundNumber,
                    demo.CurrentGameTick.Value,
                    player.PlayerName,
                    inventory
                ));
            }
        }

        var ticks = demo.CurrentDemoTick.Value;

        var reader = DemoFileReader.Create(demo, File.OpenRead(demoPath));
        await reader.ReadAllAsync();

        // Write all stats to database
        Console.WriteLine("\nWriting data to database...");
        await WriteAllStatsToDatabase(
            grenadePositions, // Dictionary of grenade positions by type
            grenadeEvents, // List of grenade events during the game
            grenadeInventory, // List of inventory snapshots
            roundTicks, // Dictionary of round information
            deathEvents, // List of death events
            playerRoundStats // List of player statistics
        );
        Console.WriteLine("Database write completed successfully.");

        // Continue with existing Excel and image generation if needed
        WriteToCsv(grenadePositions["Flashbang"], "Flashbang.csv", "Flashbang", demo);
        WriteToCsv(grenadePositions["Smoke"], "Smoke.csv", "Smoke", demo);
        WriteToCsv(grenadePositions["HE"], "HE.csv", "HE", demo);

        // Update the relevant part in the Main method where PlotGrenades is called
        string outputFileName = $"{currentMap.Name}.png";
        // Update the relevant part in the Main method where PlotGrenades is called
        string mapImagePath = $"maps/de_{currentMap.Name}.png";
        PlotGrenades(demoPath, mapImagePath); // Pass the demo path instead of constructing a filename

        if (currentMap.HasLowerLevel)
        {
            string lowerMapImagePath = $"maps/de_{currentMap.Name}_lower.png";
            PlotGrenades(demoPath, lowerMapImagePath, true);
        }
    }

    // Helper method to generate a unique filename
    private static string GetUniqueFileName(string basePath, string fileName)
    {
        // Get the file name without extension and the extension
        string nameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
        string extension = Path.GetExtension(fileName);
        string fullPath = Path.Combine(basePath, fileName);
        
        // If the file doesn't exist, return the original name
        if (!File.Exists(fullPath))
            return fileName;
        
        // If it exists, try adding numbers until we find a unique name
        int counter = 1;
        string newFileName;
        do
        {
            newFileName = $"{nameWithoutExtension}({counter}){extension}";
            fullPath = Path.Combine(basePath, newFileName);
            counter++;
        } while (File.Exists(fullPath));
        
        return newFileName;
    }

    // Helper method to ensure directory exists
    private static void EnsureDirectoryExists(string path)
    {
        string directoryPath = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directoryPath) && !Directory.Exists(directoryPath))
        {
            Directory.CreateDirectory(directoryPath);
        }
    }

    private static void PlotGrenades(string demoPath, string mapImagePath, bool lowerLevel = false)
    {
        try
        {
            Console.WriteLine($"\nProcessing {(lowerLevel ? "lower" : "upper")} level image...");

            // Get the base name from the demo file
            string baseName = GetBaseNameFromDemo(demoPath);
            
            // Construct the output filename, adding _lower suffix if needed
            string outputFileName = lowerLevel ? $"{baseName}_lower.png" : $"{baseName}.png";

            // Construct the full paths
            string wwwrootPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
            string mapsOutputPath = Path.Combine(wwwrootPath, "maps");
            
            // Ensure the maps directory exists
            EnsureDirectoryExists(mapsOutputPath);
            
            string fullOutputPath = Path.Combine(mapsOutputPath, outputFileName);
            
            // Use the map image path relative to the wwwroot directory
            string fullMapImagePath = Path.Combine(wwwrootPath, mapImagePath);

            using var baseImage = new Bitmap(fullMapImagePath);
            using var bitmap = new Bitmap(baseImage.Width, baseImage.Height, PixelFormat.Format32bppArgb);

            Console.WriteLine("Working!");
            
            using var graphics = Graphics.FromImage(bitmap);

            Console.WriteLine("Drawing grenades...");
            
            int grenadeCount = 0;
            
            // Draw the background image
            graphics.DrawImage(baseImage, 0, 0, baseImage.Width, baseImage.Height);

            var colors = new Dictionary<string, Color>
            {
                { "Smoke", Color.FromArgb(180, Color.Blue) },
                { "Molotov", Color.FromArgb(180, Color.Red) },
                { "HE", Color.FromArgb(180, Color.Green) },
                { "Flashbang", Color.FromArgb(180, Color.Yellow) }
            };

            foreach (var (grenadeType, positions) in grenadePositions)
            {
                using var brush = new SolidBrush(colors[grenadeType]);
                foreach (var (x, y, z, player, _) in positions)
                {
                    bool shouldPlot = false;
                    
                    if (currentMap.HasLowerLevel)
                    {
                        if (lowerLevel)
                        {
                            shouldPlot = z <= currentMap.LowerAltitudeMax && z >= currentMap.LowerAltitudeMin;
                        }
                        else
                        {
                            shouldPlot = z <= currentMap.DefaultAltitudeMax && z >= currentMap.DefaultAltitudeMin;
                        }
                    }
                    else
                    {
                        shouldPlot = true;
                    }

                    if (shouldPlot)
                    {
                        var imageX = ConvertToImageX(x, baseImage.Width);
                        var imageY = ConvertToImageY(y, baseImage.Height);
                        graphics.FillEllipse(brush, imageX - 4, imageY - 4, 8, 8);
                        grenadeCount++;
                        Console.WriteLine($"Plotting {grenadeType} at Z: {z} on {(lowerLevel ? "lower" : "upper")} level image");
                    }
                }
            }

            Console.WriteLine($"Total grenades plotted on {(lowerLevel ? "lower" : "upper")} level: {grenadeCount}");

            var encoderParameters = new EncoderParameters(1);
            encoderParameters.Param[0] = new EncoderParameter(Encoder.Quality, 100L);
            bitmap.Save(fullOutputPath, GetEncoder(ImageFormat.Png), encoderParameters);
            
            Console.WriteLine($"Saved {fullOutputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing image: {ex.Message}");
            throw;
        }
    }

    private static float ConvertToImageX(float gameX, int imageWidth)
    {
        // Convert from game coordinates to radar coordinates
        float normalized = (gameX - currentMap.PosX) / currentMap.Scale;
        // Convert to image coordinates (1024x1024)
        return normalized * imageWidth / 1024f;
    }

    private static float ConvertToImageY(float gameY, int imageHeight)
    {
        // Convert from game coordinates to radar coordinates, flip Y axis
        float normalized = (currentMap.PosY - gameY) / currentMap.Scale;
        // Convert to image coordinates (1024x1024)
        return normalized * imageHeight / 1024f;
    }

    private static ImageCodecInfo GetEncoder(ImageFormat format)
    {
        var codecs = ImageCodecInfo.GetImageDecoders();
        foreach (var codec in codecs)
        {
            if (codec.FormatID == format.Guid)
            {
                return codec;
            }
        }
        return null;
    }
    
    // Helper method to convert round numbers to ValidRound objects
    private static List<ValidRound> ConvertToValidRounds(List<int> roundNumbers, string demoId)
    {
        if (string.IsNullOrEmpty(demoId))
        {
            throw new ArgumentException("DemoId cannot be null or empty", nameof(demoId));
        }

        return roundNumbers.Select(roundNum => new ValidRound
        {
            RoundNumber = roundNum,
            MVP = null,
            MVPReason = null,
            IsMatchPoint = false,
            IsFinal = false,
            IsLastRoundHalf = false
        }).ToList();
    }
    
    private static async Task WritePlayerStatsToDatabase(List<PlayerStats> playerRoundStats)
    {
        // First validate and clean up the data
        CleanupRoundNumbers();
        
        Console.WriteLine($"\nWriting player stats to database. Total stats entries: {playerRoundStats.Count}");

        // Connection string - replace with your actual connection details
        string connectionString = "Server=YOUR_SERVER;Database=YOUR_DB;User Id=YOUR_USER;Password=YOUR_PASSWORD;";

        try
        {
            await using var connection = new SqlConnection(connectionString);
            await connection.OpenAsync();

            // Create the table if it doesn't exist
            string createTableSql = @"
                IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'PlayerStats')
                CREATE TABLE PlayerStats (
                    ID INT IDENTITY(1,1) PRIMARY KEY,
                    Round INT NOT NULL,
                    PlayerName NVARCHAR(100) NOT NULL,
                    SteamID NVARCHAR(50) NOT NULL,
                    Team NVARCHAR(10) NOT NULL,
                    Ping INT NOT NULL,
                    Money INT NOT NULL,
                    Kills INT NOT NULL,
                    Deaths INT NOT NULL,
                    Assists INT NOT NULL,
                    HeadshotPercentage FLOAT NOT NULL,
                    Damage INT NOT NULL,
                    CreatedAt DATETIME2 DEFAULT GETUTCDATE()
                )";

            await using (var createCmd = new SqlCommand(createTableSql, connection))
            {
                await createCmd.ExecuteNonQueryAsync();
            }

            // Sort stats for consistent ordering
            var sortedStats = playerRoundStats
                .OrderBy(x => x.Round)
                .ThenBy(x => x.PlayerName)
                .ToList();

            // Prepare batch insert
            var dataTable = new DataTable();
            dataTable.Columns.Add("Round", typeof(int));
            dataTable.Columns.Add("PlayerName", typeof(string));
            dataTable.Columns.Add("SteamID", typeof(string));
            dataTable.Columns.Add("Team", typeof(string));
            dataTable.Columns.Add("Ping", typeof(int));
            dataTable.Columns.Add("Money", typeof(int));
            dataTable.Columns.Add("Kills", typeof(int));
            dataTable.Columns.Add("Deaths", typeof(int));
            dataTable.Columns.Add("Assists", typeof(int));
            dataTable.Columns.Add("HeadshotPercentage", typeof(double));
            dataTable.Columns.Add("Damage", typeof(int));

            // Add data to DataTable
            foreach (var stat in sortedStats)
            {
                dataTable.Rows.Add(
                    stat.Round,
                    stat.PlayerName,
                    stat.SteamID,
                    stat.Team,
                    stat.Ping,
                    stat.Money,
                    stat.Kills,
                    stat.Deaths,
                    stat.Assists,
                    stat.HeadshotPercentage,
                    stat.Damage
                );
            }

            // Bulk insert the data
            using (var bulkCopy = new SqlBulkCopy(connection))
            {
                bulkCopy.DestinationTableName = "PlayerStats";
                bulkCopy.BatchSize = 1000; // Adjust based on your needs

                // Map the columns
                bulkCopy.ColumnMappings.Add("Round", "Round");
                bulkCopy.ColumnMappings.Add("PlayerName", "PlayerName");
                bulkCopy.ColumnMappings.Add("SteamID", "SteamID");
                bulkCopy.ColumnMappings.Add("Team", "Team");
                bulkCopy.ColumnMappings.Add("Ping", "Ping");
                bulkCopy.ColumnMappings.Add("Money", "Money");
                bulkCopy.ColumnMappings.Add("Kills", "Kills");
                bulkCopy.ColumnMappings.Add("Deaths", "Deaths");
                bulkCopy.ColumnMappings.Add("Assists", "Assists");
                bulkCopy.ColumnMappings.Add("HeadshotPercentage", "HeadshotPercentage");
                bulkCopy.ColumnMappings.Add("Damage", "Damage");

                // Add event handler for bulk copy notifications
                bulkCopy.NotifyAfter = 1000;
                bulkCopy.SqlRowsCopied += (sender, e) => 
                {
                    Console.WriteLine($"Copied {e.RowsCopied} rows so far...");
                };

                await bulkCopy.WriteToServerAsync(dataTable);
            }

            Console.WriteLine($"Successfully wrote {sortedStats.Count} rows to database");

            // Optional: Add an index for better query performance
            string createIndexSql = @"
                IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_PlayerStats_Round_PlayerName')
                CREATE INDEX IX_PlayerStats_Round_PlayerName ON PlayerStats (Round, PlayerName)";

            await using (var indexCmd = new SqlCommand(createIndexSql, connection))
            {
                await indexCmd.ExecuteNonQueryAsync();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error writing to database: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
            throw;
        }
    }
    
    private static async Task WriteDemoInfo(SqlConnection connection, SqlTransaction transaction, string demoPath, string mapName)
    {
        // Create demo info object with parsed data
        var demoInfo = new DemoInfo
        {
            // Use the full path as a unique identifier after normalizing it
            // Extract just the filename from the path
            DemoName = Path.GetFileName(demoPath),
            // For now, we'll use a simple hash of the demo path as the match ID
            // In a real implementation, you might want to derive this from demo naming conventions
            MatchId = Convert.ToBase64String(
                System.Security.Cryptography.SHA256.HashData(
                    System.Text.Encoding.UTF8.GetBytes(demoPath)
                )).Substring(0, 20),
            MapName = mapName,
            RecordedAt = File.GetLastWriteTime(demoPath) // Use file timestamp as recording time
        };

        // Create DataTable for bulk insert
        var dataTable = new DataTable();
        dataTable.Columns.Add("DemoName", typeof(string));
        dataTable.Columns.Add("MatchId", typeof(string));
        dataTable.Columns.Add("MapName", typeof(string));
        dataTable.Columns.Add("RecordedAt", typeof(DateTime));

        // Add the demo info to the DataTable
        dataTable.Rows.Add(
            demoInfo.DemoName,
            demoInfo.MatchId,
            demoInfo.MapName,
            demoInfo.RecordedAt
        );

        // Define column mappings
        var columnMappings = new Dictionary<string, string>
        {
            { "DemoName", "DemoName" },
            { "MatchId", "MatchId" },
            { "MapName", "MapName" },
            { "RecordedAt", "RecordedAt" }
        };

        try 
        {
            // First, check if this demo already exists
            var checkCommand = new SqlCommand(
                "SELECT COUNT(*) FROM Demos WHERE DemoName = @DemoName", 
                connection, 
                transaction);
            checkCommand.Parameters.AddWithValue("@DemoName", demoInfo.DemoName);
            
            var exists = (int)await checkCommand.ExecuteScalarAsync() > 0;

            if (!exists)
            {
                // If it doesn't exist, perform the bulk insert
                await BulkInsertData(connection, transaction, "Demos", dataTable, columnMappings);
                Console.WriteLine($"Successfully wrote demo info for {demoInfo.DemoName}");
            }
            else
            {
                Console.WriteLine($"Demo {demoInfo.DemoName} already exists in database, skipping insert");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error writing demo info to database: {ex.Message}");
            throw;
        }

        // Store the demo info for use in other methods
        currentDemoInfo = demoInfo;
    }
    
    // 3. Update the WritePlayerStatsToExcel method with debug logging
    private static async Task WritePlayerStatsToExcel(XLWorkbook workbook)
    {
        // Clean up round numbers before writing
        CleanupRoundNumbers();
        
        Console.WriteLine($"\nWriting player stats to Excel. Total stats entries: {playerRoundStats.Count}");
        // var statsSheet = workbook.Worksheets.Add("Player Stats");
        //
        // // Write headers
        // statsSheet.Cell(1, 1).Value = "Round";
        // statsSheet.Cell(1, 2).Value = "Player Name";
        // statsSheet.Cell(1, 3).Value = "Steam ID";
        // statsSheet.Cell(1, 4).Value = "Team";
        // statsSheet.Cell(1, 5).Value = "Ping";
        // statsSheet.Cell(1, 6).Value = "Money";
        // statsSheet.Cell(1, 7).Value = "Kills";
        // statsSheet.Cell(1, 8).Value = "Deaths";
        // statsSheet.Cell(1, 9).Value = "Assists";
        // statsSheet.Cell(1, 10).Value = "Headshot %";
        // statsSheet.Cell(1, 11).Value = "Damage";
        //
        // Sort by round number then player name for consistent ordering
        var sortedStats = playerRoundStats.OrderBy(x => x.Round)
                                         .ThenBy(x => x.PlayerName)
                                         .ToList();
        
        // Call the database writer instead
        await WritePlayerStatsToDatabase(playerRoundStats);
        
        //
        // var rowIndex = 2;
        // foreach (var stat in sortedStats)
        // {
        //     Console.WriteLine($"Writing row {rowIndex}: Round {stat.Round} - {stat.PlayerName}");
        //
        //     statsSheet.Cell(rowIndex, 1).Value = stat.Round;
        //     statsSheet.Cell(rowIndex, 2).Value = stat.PlayerName;
        //     statsSheet.Cell(rowIndex, 3).Value = stat.SteamID;
        //     statsSheet.Cell(rowIndex, 4).Value = stat.Team;
        //     statsSheet.Cell(rowIndex, 5).Value = stat.Ping;
        //     statsSheet.Cell(rowIndex, 6).Value = stat.Money;
        //     statsSheet.Cell(rowIndex, 7).Value = stat.Kills;
        //     statsSheet.Cell(rowIndex, 8).Value = stat.Deaths;
        //     statsSheet.Cell(rowIndex, 9).Value = stat.Assists;
        //     statsSheet.Cell(rowIndex, 10).Value = stat.HeadshotPercentage;
        //     statsSheet.Cell(rowIndex, 11).Value = stat.Damage;
        //     
        //     rowIndex++;
        // }
        //
        // Console.WriteLine($"Finished writing {rowIndex - 2} rows of player stats");
        // statsSheet.Columns().AdjustToContents();
    }

    // First, let's create a method to validate if a round is valid
    private static bool IsValidGameplayRound(int roundNum, CsDemoParser demo)
    {
        // Check if this round has actual gameplay
        if (!roundTicks.ContainsKey(roundNum))
            return false;

        var roundInfo = roundTicks[roundNum];
        
        // A valid round should have:
        // 1. A start tick
        // 2. An end tick
        // 3. At least one player on each team
        bool hasValidTicks = roundInfo.StartTick > 0 && roundInfo.EndTick > 0;
        bool hasTeams = demo.TeamTerrorist?.CSPlayerControllers.Any() == true && 
                       demo.TeamCounterTerrorist?.CSPlayerControllers.Any() == true;
                       
        return hasValidTicks && hasTeams;
    }
    
    private static void UpdateRoundEvents()
    {
        Console.WriteLine("\nStarting round number cleanup...");
        
        // First, get all unique round numbers from all data structures
        var allRounds = new HashSet<int>();
        
        // Collect rounds from each data structure
        allRounds.UnionWith(playerRoundStats.Select(x => x.Round));
        allRounds.UnionWith(grenadeEvents.Select(x => x.roundNum));
        allRounds.UnionWith(grenadeInventory.Select(x => x.roundNum));
        allRounds.UnionWith(deathEvents.Select(x => x.RoundNumber));
        allRounds.UnionWith(roundTicks.Keys);
        
        Console.WriteLine($"All round numbers found: {string.Join(", ", allRounds.OrderBy(x => x))}");
        
        // Filter out round 0 and get valid rounds
        var validRounds = allRounds.Where(r => r > 0).OrderBy(x => x).ToList();
        Console.WriteLine($"Valid gameplay rounds (excluding warmup): {string.Join(", ", validRounds)}");
        
        // Create sequential mapping starting from 1
        var sequentialMapping = new Dictionary<int, int>();
        for (int i = 0; i < validRounds.Count; i++)
        {
            sequentialMapping[validRounds[i]] = i + 1;
            Console.WriteLine($"Mapping round {validRounds[i]} to {i + 1}");
        }

        // Helper function to safely get mapped round number
        int GetMappedRound(int originalRound)
        {
            return sequentialMapping.TryGetValue(originalRound, out int mappedRound) ? mappedRound : originalRound;
        }

        // Update player stats
        playerRoundStats = playerRoundStats
            .Where(stat => stat.Round > 0)
            .Select(stat => new PlayerStats
            {
                Round = GetMappedRound(stat.Round),
                PlayerName = stat.PlayerName,
                SteamID = stat.SteamID,
                Team = stat.Team,
                Ping = stat.Ping,
                Money = stat.Money,
                Kills = stat.Kills,
                Deaths = stat.Deaths,
                Assists = stat.Assists,
                HeadshotPercentage = stat.HeadshotPercentage,
                Damage = stat.Damage
            })
            .ToList();

        // Update grenade events
        grenadeEvents = grenadeEvents
            .Where(e => e.roundNum > 0)
            .Select(evt => (GetMappedRound(evt.roundNum), evt.Value, evt.Item3, evt.Item4, evt.Weapon))
            .ToList();

        // Update grenade inventory
        grenadeInventory = grenadeInventory
            .Where(inv => inv.roundNum > 0)
            .Select(inv => (GetMappedRound(inv.roundNum), inv.Value, inv.PlayerName, inv.inventory))
            .ToList();

        // Update death events
        deathEvents = deathEvents
            .Where(d => d.RoundNumber > 0)
            .Select(d => new DeathEvent
            {
                RoundNumber = GetMappedRound(d.RoundNumber),
                Tick = d.Tick,
                AttackerName = d.AttackerName,
                AttackerTeam = d.AttackerTeam,
                VictimName = d.VictimName,
                VictimTeam = d.VictimTeam,
                Weapon = d.Weapon,
                Headshot = d.Headshot
            })
            .ToList();

        // Update round ticks
        var updatedRoundTicks = new Dictionary<int, RoundInfo>();
        foreach (var (round, info) in roundTicks.Where(kvp => kvp.Key > 0))
        {
            var newRoundNumber = GetMappedRound(round);
            updatedRoundTicks[newRoundNumber] = info;
        }
        roundTicks = updatedRoundTicks;

        // Print final distribution of rounds in each data structure
        Console.WriteLine("\nFinal round distribution after cleanup and renumbering:");
        Console.WriteLine($"Player Stats rounds: {string.Join(", ", playerRoundStats.Select(x => x.Round).Distinct().OrderBy(x => x))}");
        Console.WriteLine($"Grenade Events rounds: {string.Join(", ", grenadeEvents.Select(x => x.roundNum).Distinct().OrderBy(x => x))}");
        Console.WriteLine($"Inventory rounds: {string.Join(", ", grenadeInventory.Select(x => x.roundNum).Distinct().OrderBy(x => x))}");
        Console.WriteLine($"Death Events rounds: {string.Join(", ", deathEvents.Select(x => x.RoundNumber).Distinct().OrderBy(x => x))}");
        Console.WriteLine($"Round Ticks rounds: {string.Join(", ", roundTicks.Keys.OrderBy(x => x))}");

        // Verify all data structures have consistent round numbers
        var finalRounds = new HashSet<int>();
        finalRounds.UnionWith(playerRoundStats.Select(x => x.Round));
        finalRounds.UnionWith(grenadeEvents.Select(x => x.roundNum));
        finalRounds.UnionWith(grenadeInventory.Select(x => x.roundNum));
        finalRounds.UnionWith(deathEvents.Select(x => x.RoundNumber));
        finalRounds.UnionWith(roundTicks.Keys);

        Console.WriteLine($"\nVerification - All round numbers are now: {string.Join(", ", finalRounds.OrderBy(x => x))}");
    }
    
    private static void WriteToCsv(List<(float X, float Y, float Z, string Player, string DemoId)> lines, string outputPath, string type, CsDemoParser demo)
    {
        try
        {
            if (type != "Flashbang")
            {
                return;
            }
            
            // Before writing to Excel, update all round numbers to match player stats
            UpdateRoundEvents();

            Console.WriteLine("\n=== Debug Data Collections ===");
            Console.WriteLine($"Total grenade positions: {grenadePositions.Sum(x => x.Value.Count)}");
            foreach (var (gType, positions) in grenadePositions)
            {
                Console.WriteLine($"{gType}: {positions.Count} positions");
            }
            Console.WriteLine($"Total grenade events: {grenadeEvents.Count}");
            Console.WriteLine($"Total inventory snapshots: {grenadeInventory.Count}");
            Console.WriteLine($"Total player stats records: {playerRoundStats.Count}");
            Console.WriteLine($"Total round ticks records: {roundTicks.Count}");
            Console.WriteLine($"Total death events: {deathEvents.Count}");

            string excelPath = "grenades.xlsx";
            using var workbook = new XLWorkbook();

            var validRounds = playerRoundStats.Select(x => x.Round).Distinct().OrderBy(x => x).ToList();
            Console.WriteLine($"\nValid gameplay rounds: {string.Join(", ", validRounds)}");
            Console.WriteLine("Rounds in roundTicks: " + string.Join(", ", roundTicks.Keys.OrderBy(x => x)));

            // Write grenade positions
            foreach (var grenadeType in grenadePositions.Keys)
            {
                var worksheet = workbook.Worksheets.Add($"{grenadeType} Positions");
                
                worksheet.Cell(1, 1).Value = "Player";
                worksheet.Cell(1, 2).Value = "X";
                worksheet.Cell(1, 3).Value = "Y";
                worksheet.Cell(1, 4).Value = "Z";
                worksheet.Cell(1, 5).Value = "Level";

                var rowIndex = 2;
                foreach (var (x, y, z, player, _) in grenadePositions[grenadeType])
                {
                    string level = z <= currentMap.LowerAltitudeMax ? "Lower" : "Upper";
                    
                    worksheet.Cell(rowIndex, 1).Value = player;
                    worksheet.Cell(rowIndex, 2).Value = x;
                    worksheet.Cell(rowIndex, 3).Value = y;
                    worksheet.Cell(rowIndex, 4).Value = z;
                    worksheet.Cell(rowIndex, 5).Value = level;
                    
                    rowIndex++;
                }
                worksheet.Columns().AdjustToContents();
            }

            // Write grenade events
            var eventsSheet = workbook.Worksheets.Add("Grenade Events");
            eventsSheet.Cell(1, 1).Value = "Round";
            eventsSheet.Cell(1, 2).Value = "Tick";
            eventsSheet.Cell(1, 3).Value = "Player";
            eventsSheet.Cell(1, 4).Value = "Action";
            eventsSheet.Cell(1, 5).Value = "Grenade Type";

            var eventRowIndex = 2;
            foreach (var evt in grenadeEvents.Where(e => validRounds.Contains(e.roundNum)))
            {
                if (!roundTicks.TryGetValue(evt.roundNum, out var roundInfo))
                {
                    Console.WriteLine($"Warning: No round info found for round {evt.roundNum}");
                    continue;
                }

                eventsSheet.Cell(eventRowIndex, 1).Value = evt.roundNum;
                eventsSheet.Cell(eventRowIndex, 2).Value = evt.Value;
                eventsSheet.Cell(eventRowIndex, 3).Value = evt.Item3;
                eventsSheet.Cell(eventRowIndex, 4).Value = evt.Item4;
                eventsSheet.Cell(eventRowIndex, 5).Value = evt.Weapon;
                eventRowIndex++;
            }
            eventsSheet.Columns().AdjustToContents();

            // Write inventory snapshots
            var inventorySheet = workbook.Worksheets.Add("Inventory Snapshots");
            inventorySheet.Cell(1, 1).Value = "Round";
            inventorySheet.Cell(1, 2).Value = "Tick";
            inventorySheet.Cell(1, 3).Value = "Player";
            inventorySheet.Cell(1, 4).Value = "Inventory";

            var invRowIndex = 2;
            foreach (var inv in grenadeInventory.Where(s => validRounds.Contains(s.roundNum)))
            {
                if (!roundTicks.TryGetValue(inv.roundNum, out var roundInfo))
                {
                    Console.WriteLine($"Warning: No round info found for round {inv.roundNum}");
                    continue;
                }

                inventorySheet.Cell(invRowIndex, 1).Value = inv.roundNum;
                inventorySheet.Cell(invRowIndex, 2).Value = inv.Value;
                inventorySheet.Cell(invRowIndex, 3).Value = inv.PlayerName;
                inventorySheet.Cell(invRowIndex, 4).Value = string.Join(", ", 
                    inv.inventory.Select(x => $"{x.GrenadeType} x {x.Count}"));
                invRowIndex++;
            }
            inventorySheet.Columns().AdjustToContents();

            // Write rounds data
            var roundsSheet = workbook.Worksheets.Add("Rounds");
            roundsSheet.Cell(1, 1).Value = "Round";
            roundsSheet.Cell(1, 2).Value = "Start Tick";
            roundsSheet.Cell(1, 3).Value = "Freeze End Tick";
            roundsSheet.Cell(1, 4).Value = "End Tick";
            roundsSheet.Cell(1, 5).Value = "Round Duration";
            roundsSheet.Cell(1, 6).Value = "Freeze Time";
            roundsSheet.Cell(1, 7).Value = "MVP";
            roundsSheet.Cell(1, 8).Value = "MVP Reason";
            roundsSheet.Cell(1, 9).Value = "Match Point";
            roundsSheet.Cell(1, 10).Value = "Final Round";
            roundsSheet.Cell(1, 11).Value = "Last Round in Half";

            var roundRowIndex = 2;
            foreach (var (round, info) in roundTicks.OrderBy(x => x.Key))
            {
                if (!validRounds.Contains(round))
                    continue;

                if (info.EndTick > 0)
                {
                    roundsSheet.Cell(roundRowIndex, 1).Value = round;
                    roundsSheet.Cell(roundRowIndex, 2).Value = info.StartTick;
                    roundsSheet.Cell(roundRowIndex, 3).Value = info.FreezeEndTick;
                    roundsSheet.Cell(roundRowIndex, 4).Value = info.EndTick;
                    roundsSheet.Cell(roundRowIndex, 5).Value = info.EndTick - info.StartTick;
                    roundsSheet.Cell(roundRowIndex, 6).Value = info.FreezeEndTick > 0 ? 
                        info.FreezeEndTick - info.StartTick : 0;
                    roundsSheet.Cell(roundRowIndex, 7).Value = string.IsNullOrEmpty(info.MVP) ? "Unknown" : info.MVP;
                    roundsSheet.Cell(roundRowIndex, 8).Value = info.MvpReason;
                    roundsSheet.Cell(roundRowIndex, 9).Value = info.IsMatchPoint ? "Yes" : "No";
                    roundsSheet.Cell(roundRowIndex, 10).Value = info.IsFinal ? "Yes" : "No";
                    roundsSheet.Cell(roundRowIndex, 11).Value = info.IsLastRoundHalf ? "Yes" : "No";
                    
                    roundRowIndex++;
                }
            }
            roundsSheet.Columns().AdjustToContents();

            // Write death events
            var deathSheet = workbook.Worksheets.Add("Deaths");
            deathSheet.Cell(1, 1).Value = "Round";
            deathSheet.Cell(1, 2).Value = "Tick";
            deathSheet.Cell(1, 3).Value = "Time in Round";
            deathSheet.Cell(1, 4).Value = "Attacker";
            deathSheet.Cell(1, 5).Value = "Attacker Team";
            deathSheet.Cell(1, 6).Value = "Victim";
            deathSheet.Cell(1, 7).Value = "Victim Team";
            deathSheet.Cell(1, 8).Value = "Weapon";
            deathSheet.Cell(1, 9).Value = "Headshot";

            var deathRowIndex = 2;
            foreach (var death in deathEvents.Where(d => validRounds.Contains(d.RoundNumber)))
            {
                if (!roundTicks.TryGetValue(death.RoundNumber, out var roundInfo))
                {
                    Console.WriteLine($"Warning: No round info found for round {death.RoundNumber}");
                    continue;
                }

                var timeInRound = roundInfo.StartTick > 0 ? 
                    (death.Tick - roundInfo.StartTick) / CsDemoParser.TickRate : 0;

                deathSheet.Cell(deathRowIndex, 1).Value = death.RoundNumber;
                deathSheet.Cell(deathRowIndex, 2).Value = death.Tick;
                deathSheet.Cell(deathRowIndex, 3).Value = timeInRound;
                deathSheet.Cell(deathRowIndex, 4).Value = death.AttackerName;
                deathSheet.Cell(deathRowIndex, 5).Value = death.AttackerTeam;
                deathSheet.Cell(deathRowIndex, 6).Value = death.VictimName;
                deathSheet.Cell(deathRowIndex, 7).Value = death.VictimTeam;
                deathSheet.Cell(deathRowIndex, 8).Value = death.Weapon;
                deathSheet.Cell(deathRowIndex, 9).Value = death.Headshot ? "Yes" : "No";
        
                deathRowIndex++;
            }
            deathSheet.Columns().AdjustToContents();

            // Write player stats
            WritePlayerStatsToExcel(workbook);

            var directory = Path.GetDirectoryName(excelPath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.SaveAs(excelPath);
            Console.WriteLine($"\nSuccessfully wrote all data to {excelPath}");
            
            Console.WriteLine("\nSummary of written data:");
            Console.WriteLine($"Total valid rounds: {validRounds.Count}");
            Console.WriteLine($"Round numbers: {string.Join(", ", validRounds)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error writing to Excel file: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
            throw;
        }
    }
        
    private static readonly string[] PlayerColours =
    {
        "#FF0000",
        "#FF7F00",
        "#FFD700",
        "#7FFF00",
        "#00FF00",
        "#00FF7F",
        "#00FFFF",
        "#007FFF",
        "#0000FF",
        "#7F00FF",
        "#FF00FF",
        "#FF007F",
    };
    
    private static void MarkupPlayerName(CBasePlayerController? player)
    {
        if (player == null)
        {
            AnsiConsole.Markup("[grey](unknown)[/]");
            return;
        }

        AnsiConsole.Markup($"[{PlayerColours[player.EntityIndex.Value % PlayerColours.Length]}]{player.PlayerName}[/]");
    }
    
    public static async Task WriteAllStatsToDatabase(
    Dictionary<string, List<(float X, float Y, float Z, string Player, string DemoId)>> grenadePositions,
    List<(int roundNum, uint Value, string Item3, string Item4, string Weapon)> grenadeEvents,
    List<(int roundNum, uint Value, string PlayerName, List<(string GrenadeType, int Count)> inventory)> grenadeInventory,
    Dictionary<int, RoundInfo> roundTicks,
    List<DeathEvent> deathEvents,
    List<PlayerStats> playerRoundStats)
    {
        try
        {
            // First validate and clean up the data
            UpdateRoundEvents();

            // Get valid rounds from playerRoundStats after cleanup
            // We use playerRoundStats because it's our most reliable source of round information
            var validRounds = playerRoundStats
                .Select(x => x.Round)
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            // Log initial statistics for verification
            Console.WriteLine("\n=== Debug Data Collections ===");
            Console.WriteLine($"Total grenade positions: {grenadePositions.Sum(x => x.Value.Count)}");
            Console.WriteLine($"Total grenade events: {grenadeEvents.Count}");
            Console.WriteLine($"Total inventory snapshots: {grenadeInventory.Count}");
            Console.WriteLine($"Total player stats records: {playerRoundStats.Count}");
            Console.WriteLine($"Total round ticks records: {roundTicks.Count}");
            Console.WriteLine($"Total death events: {deathEvents.Count}");
            Console.WriteLine($"\nValid gameplay rounds: {string.Join(", ", validRounds)}");

            await using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            await CreateDatabaseSchema(connection);

            await using var transaction = connection.BeginTransaction();
            try
            {
                await WriteDemoInfo(connection, transaction, currentDemoInfo.DemoName, currentDemoInfo.MapName);

                // Then write the rounds and other data...
                await WriteRoundData(connection, transaction, roundTicks, validRounds);
                
                // Then write the dependent data
                await WriteGrenadePositions(connection, transaction, grenadePositions);
                await WriteGrenadeEvents(connection, transaction, grenadeEvents, validRounds);
                await WriteInventorySnapshots(connection, transaction, grenadeInventory, validRounds);
                await WriteDeathEvents(connection, transaction, deathEvents, validRounds, roundTicks);
                await WritePlayerStats(connection, transaction, playerRoundStats);

                await transaction.CommitAsync();
                Console.WriteLine("\nSuccessfully committed all data to database");
            }
            catch (Exception)
            {
                Console.WriteLine("\nError occurred, rolling back transaction");
                await transaction.RollbackAsync();
                throw;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"\nError writing to database: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
            throw;
        }
    }

    private static async Task WriteGrenadePositions(SqlConnection connection, SqlTransaction transaction,
        Dictionary<string, List<(float X, float Y, float Z, string Player, string DemoId)>> grenadePositions)
    {
        if (string.IsNullOrEmpty(currentDemoInfo?.DemoName))
        {
            throw new InvalidOperationException("DemoName must be available before writing grenade positions");
        }

        var dataTable = new DataTable();
        dataTable.Columns.Add("DemoName", typeof(string));      // Add DemoName first
        dataTable.Columns.Add("GrenadeType", typeof(string));
        dataTable.Columns.Add("Player", typeof(string));
        dataTable.Columns.Add("X", typeof(float));
        dataTable.Columns.Add("Y", typeof(float));
        dataTable.Columns.Add("Z", typeof(float));
        dataTable.Columns.Add("Level", typeof(string));

        Console.WriteLine("\nPreparing to write grenade positions");
        Console.WriteLine($"Total grenade types: {grenadePositions.Count}");

        foreach (var (grenadeType, positions) in grenadePositions)
        {
            Console.WriteLine($"Processing {positions.Count} {grenadeType} positions");
        
            foreach (var (x, y, z, player, _) in positions)
            {
                string level = z <= currentMap.LowerAltitudeMax ? "Lower" : "Upper";
                dataTable.Rows.Add(
                    currentDemoInfo.DemoName,  // Add DemoName for each record
                    grenadeType,
                    player,
                    x,
                    y,
                    z,
                    level
                );
            }
        }

        var columnMappings = new Dictionary<string, string>
        {
            { "DemoName", "DemoName" },
            { "GrenadeType", "GrenadeType" },
            { "Player", "Player" },
            { "X", "X" },
            { "Y", "Y" },
            { "Z", "Z" },
            { "Level", "Level" }
        };

        await BulkInsertData(connection, transaction, "GrenadePositions", dataTable, columnMappings);
        Console.WriteLine($"Successfully wrote {dataTable.Rows.Count} grenade positions");
    }

    private static async Task WriteGrenadeEvents(SqlConnection connection, SqlTransaction transaction,
        List<(int roundNum, uint Value, string Item3, string Item4, string Weapon)> grenadeEvents,
        List<int> validRounds)
    {
        if (string.IsNullOrEmpty(currentDemoInfo?.DemoName))
        {
            throw new InvalidOperationException("DemoName must be available before writing grenade events");
        }

        var dataTable = new DataTable();
        dataTable.Columns.Add("DemoName", typeof(string));      // Add DemoName first
        dataTable.Columns.Add("RoundNumber", typeof(int));
        dataTable.Columns.Add("Tick", typeof(int));
        dataTable.Columns.Add("Player", typeof(string));
        dataTable.Columns.Add("Action", typeof(string));
        dataTable.Columns.Add("GrenadeType", typeof(string));

        Console.WriteLine($"\nPreparing to write grenade events");
        Console.WriteLine($"Total events to process: {grenadeEvents.Count}");
        Console.WriteLine($"Valid rounds: {string.Join(", ", validRounds)}");

        foreach (var evt in grenadeEvents.Where(e => validRounds.Contains(e.roundNum)))
        {
            string normalizedGrenadeType = NormalizeGrenadeType(evt.Weapon);
            dataTable.Rows.Add(
                currentDemoInfo.DemoName,  // Add DemoName for each record
                evt.roundNum,
                evt.Value,
                evt.Item3,
                evt.Item4,
                normalizedGrenadeType
            );
        }

        var columnMappings = new Dictionary<string, string>
        {
            { "DemoName", "DemoName" },
            { "RoundNumber", "RoundNumber" },
            { "Tick", "Tick" },
            { "Player", "Player" },
            { "Action", "Action" },
            { "GrenadeType", "GrenadeType" }
        };

        await BulkInsertData(connection, transaction, "GrenadeEvents", dataTable, columnMappings);
        Console.WriteLine($"Successfully wrote {dataTable.Rows.Count} grenade events");
    }

    // Add this helper method to normalize grenade types
    private static string NormalizeGrenadeType(string weaponType)
    {
        // Remove the 'C' prefix and 'Grenade' suffix if present
        string normalized = weaponType
            .Replace("C", "")           // Remove C prefix
            .Replace("Grenade", "")     // Remove Grenade suffix
            .Trim();                    // Remove any extra spaces

        // If it's a comma-separated list, process each type
        if (normalized.Contains(","))
        {
            var types = normalized
                .Split(',')
                .Select(t => t.Trim())
                .Select(t => t
                    .Replace("C", "")
                    .Replace("Grenade", "")
                    .Trim())
                .Where(t => !string.IsNullOrEmpty(t));

            return string.Join(", ", types);
        }

        return normalized;
    }

    private static async Task WriteInventorySnapshots(SqlConnection connection, SqlTransaction transaction,
        List<(int roundNum, uint Value, string PlayerName, List<(string GrenadeType, int Count)> inventory)> grenadeInventory,
        List<int> validRounds)
    {
        if (string.IsNullOrEmpty(currentDemoInfo?.DemoName))
        {
            throw new InvalidOperationException("DemoName must be available before writing inventory snapshots");
        }

        var dataTable = new DataTable();
        dataTable.Columns.Add("DemoName", typeof(string));      // Add DemoName first
        dataTable.Columns.Add("RoundNumber", typeof(int));
        dataTable.Columns.Add("Tick", typeof(int));
        dataTable.Columns.Add("Player", typeof(string));
        dataTable.Columns.Add("Inventory", typeof(string));

        Console.WriteLine($"\nPreparing to write inventory snapshots");
        Console.WriteLine($"Total snapshots to process: {grenadeInventory.Count}");
        Console.WriteLine($"Valid rounds: {string.Join(", ", validRounds)}");

        foreach (var inv in grenadeInventory.Where(s => validRounds.Contains(s.roundNum)))
        {
            string inventoryStr = string.Join(", ", 
                inv.inventory.Select(x => $"{x.GrenadeType} x {x.Count}"));

            dataTable.Rows.Add(
                currentDemoInfo.DemoName,  // Add DemoName for each record
                inv.roundNum,
                inv.Value,
                inv.PlayerName,
                inventoryStr
            );
        }

        var columnMappings = new Dictionary<string, string>
        {
            { "DemoName", "DemoName" },
            { "RoundNumber", "RoundNumber" },
            { "Tick", "Tick" },
            { "Player", "Player" },
            { "Inventory", "Inventory" }
        };

        await BulkInsertData(connection, transaction, "InventorySnapshots", dataTable, columnMappings);
        Console.WriteLine($"Successfully wrote {dataTable.Rows.Count} inventory snapshots");
    }
    
    // Add this helper method to extract the base name from the demo file
    private static string GetBaseNameFromDemo(string demoPath)
    {
        // Get the filename without the path and extension
        string demoFileName = Path.GetFileNameWithoutExtension(demoPath);
    
        // Return the name, which will preserve any numbering in parentheses
        return demoFileName;
    }

    // Now let's fix the WriteRoundData method
    private static async Task WriteRoundData(SqlConnection connection, SqlTransaction transaction, 
        Dictionary<int, RoundInfo> roundTicks, List<int> roundNumbers)
    {
        // First, filter out warmup rounds
        var validRoundNumbers = roundNumbers.Where(r => r > 0).ToList();
        
        Console.WriteLine($"\nProcessing rounds for database write:");
        Console.WriteLine($"Original round numbers: {string.Join(", ", roundNumbers)}");
        Console.WriteLine($"Filtered round numbers (excluding warmup/knife): {string.Join(", ", validRoundNumbers)}");

        if (!validRoundNumbers.Any())
        {
            throw new InvalidOperationException("No valid round numbers found after filtering warmup and knife rounds");
        }

        // Check what rounds already exist in the database for this demo
        var existingRounds = new HashSet<int>();
        var checkExistingCmd = new SqlCommand(
            "SELECT RoundNumber FROM Rounds WHERE DemoName = @DemoName",
            connection,
            transaction);
        
        checkExistingCmd.Parameters.AddWithValue("@DemoName", currentDemoInfo.DemoName);

        using (var reader = await checkExistingCmd.ExecuteReaderAsync())
        {
            while (await reader.ReadAsync())
            {
                existingRounds.Add(reader.GetInt32(0));
            }
        }

        Console.WriteLine($"Found {existingRounds.Count} existing rounds in database: {string.Join(", ", existingRounds)}");

        // Filter out rounds that already exist
        var newRoundNumbers = validRoundNumbers.Where(r => !existingRounds.Contains(r)).ToList();
        Console.WriteLine($"New rounds to process: {string.Join(", ", newRoundNumbers)}");

        if (!newRoundNumbers.Any())
        {
            Console.WriteLine("No new rounds to insert - all rounds already exist in database");
            return;
        }

        // Create DataTable matching the exact database schema
        var roundsTable = new DataTable();
        roundsTable.Columns.Add("DemoName", typeof(string));
        roundsTable.Columns.Add("RoundNumber", typeof(int));
        roundsTable.Columns.Add("StartTick", typeof(int));
        roundsTable.Columns.Add("FreezeEndTick", typeof(int));
        roundsTable.Columns.Add("EndTick", typeof(int));
        roundsTable.Columns.Add("MVP", typeof(string));
        roundsTable.Columns.Add("MVPReason", typeof(int));
        roundsTable.Columns.Add("IsMatchPoint", typeof(bool));
        roundsTable.Columns.Add("IsFinal", typeof(bool));
        roundsTable.Columns.Add("IsLastRoundHalf", typeof(bool));

        // Debug logging
        Console.WriteLine("\nDataTable columns:");
        foreach (DataColumn col in roundsTable.Columns)
        {
            Console.WriteLine($"Column: {col.ColumnName}, Type: {col.DataType}");
        }

        // Populate the DataTable
        foreach (var round in newRoundNumbers)
        {
            if (!roundTicks.TryGetValue(round, out var roundTick))
            {
                Console.WriteLine($"Warning: Missing tick data for round {round} during table population");
                continue;
            }

            var row = roundsTable.NewRow();
            row["DemoName"] = currentDemoInfo.DemoName;
            row["RoundNumber"] = round;
            row["StartTick"] = roundTick.StartTick;
            row["FreezeEndTick"] = roundTick.FreezeEndTick;
            row["EndTick"] = roundTick.EndTick;
            row["MVP"] = (object)roundTick.MVP ?? DBNull.Value;
            row["MVPReason"] = roundTick.MvpReason != 0 ? (object)roundTick.MvpReason : DBNull.Value;
            row["IsMatchPoint"] = roundTick.IsMatchPoint;
            row["IsFinal"] = roundTick.IsFinal;
            row["IsLastRoundHalf"] = roundTick.IsLastRoundHalf;

            roundsTable.Rows.Add(row);
        }

        // Create mappings that match exactly with both DataTable and database columns
        var columnMappings = new Dictionary<string, string>
        {
            { "DemoName", "DemoName" },
            { "RoundNumber", "RoundNumber" },
            { "StartTick", "StartTick" },
            { "FreezeEndTick", "FreezeEndTick" },
            { "EndTick", "EndTick" },
            { "MVP", "MVP" },
            { "MVPReason", "MVPReason" },
            { "IsMatchPoint", "IsMatchPoint" },
            { "IsFinal", "IsFinal" },
            { "IsLastRoundHalf", "IsLastRoundHalf" }
        };

        try 
        {
            if (roundsTable.Rows.Count > 0)
            {
                // Final verification for duplicates
                var duplicateCheck = roundsTable.AsEnumerable()
                    .GroupBy(row => new { 
                        DemoName = row.Field<string>("DemoName"), 
                        RoundNumber = row.Field<int>("RoundNumber") 
                    })
                    .Where(g => g.Count() > 1)
                    .ToList();

                if (duplicateCheck.Any())
                {
                    throw new InvalidOperationException(
                        $"Found {duplicateCheck.Count} duplicate DemoName/RoundNumber combinations!");
                }

                Console.WriteLine($"Attempting to write {roundsTable.Rows.Count} rounds to database...");
                
                // Debug log the first row's data
                if (roundsTable.Rows.Count > 0)
                {
                    var firstRow = roundsTable.Rows[0];
                    Console.WriteLine("\nFirst row data:");
                    foreach (DataColumn col in roundsTable.Columns)
                    {
                        Console.WriteLine($"{col.ColumnName}: {firstRow[col]}");
                    }
                }

                await BulkInsertData(connection, transaction, "Rounds", roundsTable, columnMappings);
                Console.WriteLine("Successfully wrote rounds to database");
            }
        }
        catch (SqlException ex)
        {
            Console.WriteLine($"SQL Error writing rounds: {ex.Message}");
            if (ex.Number == 515)  // NULL insert violation
            {
                Console.WriteLine("Failed to insert rounds: Required data is missing");
            }
            throw;
        }
    }
    
    // Helper method to determine which round has more complete data
    private static bool IsRoundMoreComplete(ValidRound newRound, ValidRound existingRound)
    {
        int newRoundScore = 0;
        int existingRoundScore = 0;

        // Add points for each piece of data that exists
        if (!string.IsNullOrEmpty(newRound.MVP)) newRoundScore++;
        if (newRound.MVPReason.HasValue) newRoundScore++;
        if (newRound.IsMatchPoint) newRoundScore++;
        if (newRound.IsFinal) newRoundScore++;
        if (newRound.IsLastRoundHalf) newRoundScore++;

        if (!string.IsNullOrEmpty(existingRound.MVP)) existingRoundScore++;
        if (existingRound.MVPReason.HasValue) existingRoundScore++;
        if (existingRound.IsMatchPoint) existingRoundScore++;
        if (existingRound.IsFinal) existingRoundScore++;
        if (existingRound.IsLastRoundHalf) existingRoundScore++;

        return newRoundScore > existingRoundScore;
    }

    private static async Task WriteDeathEvents(SqlConnection connection, SqlTransaction transaction,
    List<DeathEvent> deathEvents, List<int> validRounds, Dictionary<int, RoundInfo> roundTicks)
    {
        // First, ensure we have a valid DemoName from currentDemoInfo
        if (string.IsNullOrEmpty(currentDemoInfo?.DemoName))
        {
            throw new InvalidOperationException("DemoName must be available before writing death events");
        }

        // Create DataTable with all required columns, including DemoName
        var dataTable = new DataTable();
        dataTable.Columns.Add("DemoName", typeof(string));      // Add DemoName first
        dataTable.Columns.Add("RoundNumber", typeof(int));
        dataTable.Columns.Add("Tick", typeof(int));
        dataTable.Columns.Add("TimeInRound", typeof(float));
        dataTable.Columns.Add("AttackerName", typeof(string));
        dataTable.Columns.Add("AttackerTeam", typeof(string));
        dataTable.Columns.Add("VictimName", typeof(string));
        dataTable.Columns.Add("VictimTeam", typeof(string));
        dataTable.Columns.Add("Weapon", typeof(string));
        dataTable.Columns.Add("Headshot", typeof(bool));

        // Add diagnostic logging
        Console.WriteLine($"\nPreparing to write death events");
        Console.WriteLine($"Total death events to process: {deathEvents.Count}");
        Console.WriteLine($"Valid rounds: {string.Join(", ", validRounds)}");
        Console.WriteLine($"Current Demo Name: {currentDemoInfo.DemoName}");

        // Populate the DataTable
        foreach (var death in deathEvents.Where(d => validRounds.Contains(d.RoundNumber)))
        {
            // Calculate time in round if we have round info
            float timeInRound = 0;
            if (roundTicks.TryGetValue(death.RoundNumber, out var roundInfo) && roundInfo.StartTick > 0)
            {
                timeInRound = (death.Tick - roundInfo.StartTick) / CsDemoParser.TickRate;
            }

            // Add all the data including DemoName
            dataTable.Rows.Add(
                currentDemoInfo.DemoName,  // DemoName - Required, non-null value
                death.RoundNumber,         // RoundNumber
                death.Tick,                // Tick
                timeInRound,               // TimeInRound
                death.AttackerName,        // AttackerName
                death.AttackerTeam,        // AttackerTeam
                death.VictimName,          // VictimName
                death.VictimTeam,          // VictimTeam
                death.Weapon,              // Weapon
                death.Headshot             // Headshot
            );
        }

        // Log a sample row for verification
        if (dataTable.Rows.Count > 0)
        {
            Console.WriteLine("\nSample row data:");
            var sampleRow = dataTable.Rows[0];
            foreach (DataColumn column in dataTable.Columns)
            {
                Console.WriteLine($"{column.ColumnName}: {sampleRow[column]}");
            }
        }

        // Define column mappings including DemoName
        var columnMappings = new Dictionary<string, string>
        {
            { "DemoName", "DemoName" },        // Add DemoName mapping
            { "RoundNumber", "RoundNumber" },
            { "Tick", "Tick" },
            { "TimeInRound", "TimeInRound" },
            { "AttackerName", "AttackerName" },
            { "AttackerTeam", "AttackerTeam" },
            { "VictimName", "VictimName" },
            { "VictimTeam", "VictimTeam" },
            { "Weapon", "Weapon" },
            { "Headshot", "Headshot" }
        };

        try 
        {
            await BulkInsertData(connection, transaction, "DeathEvents", dataTable, columnMappings);
            Console.WriteLine($"Successfully wrote {dataTable.Rows.Count} death events to database");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during death events bulk insert:");
            Console.WriteLine($"Total rows attempted: {dataTable.Rows.Count}");
            Console.WriteLine($"Error: {ex.Message}");
            throw;
        }
    }

    private static async Task WritePlayerStats(SqlConnection connection, SqlTransaction transaction,
    List<PlayerStats> playerRoundStats)
    {
        // Validate that we have a demo name before proceeding
        if (string.IsNullOrEmpty(currentDemoInfo?.DemoName))
        {
            throw new InvalidOperationException("DemoName must be available before writing player stats");
        }

        var dataTable = new DataTable();
        dataTable.Columns.Add("DemoName", typeof(string));      // Add DemoName first
        dataTable.Columns.Add("RoundNumber", typeof(int));
        dataTable.Columns.Add("PlayerName", typeof(string));
        dataTable.Columns.Add("SteamID", typeof(string));
        dataTable.Columns.Add("Team", typeof(string));
        dataTable.Columns.Add("Ping", typeof(int));
        dataTable.Columns.Add("Money", typeof(int));
        dataTable.Columns.Add("Kills", typeof(int));
        dataTable.Columns.Add("Deaths", typeof(int));
        dataTable.Columns.Add("Assists", typeof(int));
        dataTable.Columns.Add("HeadshotPercentage", typeof(float));
        dataTable.Columns.Add("Damage", typeof(int));

        Console.WriteLine($"\nPreparing to write player stats");
        Console.WriteLine($"Total stats records to process: {playerRoundStats.Count}");

        // Sort stats for consistent ordering and populate DataTable
        var sortedStats = playerRoundStats
            .OrderBy(x => x.Round)
            .ThenBy(x => x.PlayerName)
            .ToList();

        foreach (var stat in sortedStats)
        {
            dataTable.Rows.Add(
                currentDemoInfo.DemoName,   // Add DemoName for each record
                stat.Round,
                stat.PlayerName,
                stat.SteamID,
                stat.Team,
                stat.Ping,
                stat.Money,
                stat.Kills,
                stat.Deaths,
                stat.Assists,
                stat.HeadshotPercentage,
                stat.Damage
            );
        }

        // Log sample data for verification
        if (dataTable.Rows.Count > 0)
        {
            Console.WriteLine("\nSample player stats row:");
            var sampleRow = dataTable.Rows[0];
            foreach (DataColumn column in dataTable.Columns)
            {
                Console.WriteLine($"{column.ColumnName}: {sampleRow[column]}");
            }
        }

        var columnMappings = new Dictionary<string, string>
        {
            { "DemoName", "DemoName" },           // Add DemoName mapping
            { "RoundNumber", "RoundNumber" },
            { "PlayerName", "PlayerName" },
            { "SteamID", "SteamID" },
            { "Team", "Team" },
            { "Ping", "Ping" },
            { "Money", "Money" },
            { "Kills", "Kills" },
            { "Deaths", "Deaths" },
            { "Assists", "Assists" },
            { "HeadshotPercentage", "HeadshotPercentage" },
            { "Damage", "Damage" }
        };

        await BulkInsertData(connection, transaction, "PlayerStats", dataTable, columnMappings);
        Console.WriteLine($"Successfully wrote {dataTable.Rows.Count} player stats records");
    }

    private static async Task CreateDatabaseSchema(SqlConnection connection)
    {
        string createTablesSql = @"
            -- Create Demos table (parent table) first
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Demos')
            BEGIN
                CREATE TABLE Demos (
                    ID INT IDENTITY(1,1) PRIMARY KEY,
                    DemoName NVARCHAR(200) NOT NULL,
                    MatchId NVARCHAR(100) NOT NULL,
                    MapName NVARCHAR(50) NOT NULL,
                    RecordedAt DATETIME2 NOT NULL,
                    CreatedAt DATETIME2 DEFAULT GETUTCDATE(),
                    CONSTRAINT UQ_DemoName UNIQUE (DemoName)
                );
            END;

            -- Create Rounds table before other dependent tables
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Rounds')
            BEGIN
                CREATE TABLE Rounds (
                    ID INT IDENTITY(1,1) PRIMARY KEY,
                    DemoName NVARCHAR(200) NOT NULL,
                    RoundNumber INT NOT NULL,
                    StartTick INT NOT NULL,
                    FreezeEndTick INT NOT NULL,
                    EndTick INT NOT NULL,
                    MVP NVARCHAR(100) NULL,
                    MVPReason INT NULL,
                    IsMatchPoint BIT NOT NULL DEFAULT 0,
                    IsFinal BIT NOT NULL DEFAULT 0,
                    IsLastRoundHalf BIT NOT NULL DEFAULT 0,
                    CreatedAt DATETIME2 DEFAULT GETUTCDATE(),
                    CONSTRAINT FK_Rounds_Demos 
                        FOREIGN KEY (DemoName) REFERENCES Demos(DemoName),
                    CONSTRAINT UQ_Demo_Round UNIQUE (DemoName, RoundNumber)
                );
            END;

            -- Create GrenadePositions table with proper foreign key
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'GrenadePositions')
            BEGIN
                CREATE TABLE GrenadePositions (
                    ID INT IDENTITY(1,1) PRIMARY KEY,
                    DemoName NVARCHAR(200) NOT NULL,
                    GrenadeType NVARCHAR(200) NOT NULL,
                    Player NVARCHAR(100) NOT NULL,
                    X FLOAT NOT NULL,
                    Y FLOAT NOT NULL,
                    Z FLOAT NOT NULL,
                    Level NVARCHAR(20) NOT NULL,
                    CreatedAt DATETIME2 DEFAULT GETUTCDATE(),
                    CONSTRAINT FK_GrenadePositions_Demos 
                        FOREIGN KEY (DemoName) REFERENCES Demos(DemoName)
                );
            END;

            -- Create GrenadeEvents table with proper foreign keys
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'GrenadeEvents')
            BEGIN
                CREATE TABLE GrenadeEvents (
                    ID INT IDENTITY(1,1) PRIMARY KEY,
                    DemoName NVARCHAR(200) NOT NULL,
                    RoundNumber INT NOT NULL,
                    Tick INT NOT NULL,
                    Player NVARCHAR(100) NOT NULL,
                    Action NVARCHAR(50) NOT NULL,
                    GrenadeType NVARCHAR(200) NOT NULL,
                    CreatedAt DATETIME2 DEFAULT GETUTCDATE(),
                    CONSTRAINT FK_GrenadeEvents_Demos 
                        FOREIGN KEY (DemoName) REFERENCES Demos(DemoName),
                    CONSTRAINT FK_GrenadeEvents_Rounds 
                        FOREIGN KEY (DemoName, RoundNumber) 
                        REFERENCES Rounds(DemoName, RoundNumber)
                );
            END;

            -- Create InventorySnapshots table with proper foreign keys
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'InventorySnapshots')
            BEGIN
                CREATE TABLE InventorySnapshots (
                    ID INT IDENTITY(1,1) PRIMARY KEY,
                    DemoName NVARCHAR(200) NOT NULL,
                    RoundNumber INT NOT NULL,
                    Tick INT NOT NULL,
                    Player NVARCHAR(100) NOT NULL,
                    Inventory NVARCHAR(MAX) NOT NULL,
                    CreatedAt DATETIME2 DEFAULT GETUTCDATE(),
                    CONSTRAINT FK_InventorySnapshots_Demos 
                        FOREIGN KEY (DemoName) REFERENCES Demos(DemoName),
                    CONSTRAINT FK_InventorySnapshots_Rounds 
                        FOREIGN KEY (DemoName, RoundNumber) 
                        REFERENCES Rounds(DemoName, RoundNumber)
                );
            END;

            -- Create DeathEvents table with proper foreign keys
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'DeathEvents')
            BEGIN
                CREATE TABLE DeathEvents (
                    ID INT IDENTITY(1,1) PRIMARY KEY,
                    DemoName NVARCHAR(200) NOT NULL,
                    RoundNumber INT NOT NULL,
                    Tick INT NOT NULL,
                    TimeInRound FLOAT NOT NULL,
                    AttackerName NVARCHAR(100) NOT NULL,
                    AttackerTeam NVARCHAR(10) NOT NULL,
                    VictimName NVARCHAR(100) NOT NULL,
                    VictimTeam NVARCHAR(10) NOT NULL,
                    Weapon NVARCHAR(50) NOT NULL,
                    Headshot BIT NOT NULL,
                    CreatedAt DATETIME2 DEFAULT GETUTCDATE(),
                    CONSTRAINT FK_DeathEvents_Demos 
                        FOREIGN KEY (DemoName) REFERENCES Demos(DemoName),
                    CONSTRAINT FK_DeathEvents_Rounds 
                        FOREIGN KEY (DemoName, RoundNumber) 
                        REFERENCES Rounds(DemoName, RoundNumber)
                );
            END;

            -- Create PlayerStats table with proper foreign keys
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'PlayerStats')
            BEGIN
                CREATE TABLE PlayerStats (
                    ID INT IDENTITY(1,1) PRIMARY KEY,
                    DemoName NVARCHAR(200) NOT NULL,
                    RoundNumber INT NOT NULL,
                    PlayerName NVARCHAR(100) NOT NULL,
                    SteamID NVARCHAR(50) NOT NULL,
                    Team NVARCHAR(10) NOT NULL,
                    Ping INT NOT NULL,
                    Money INT NOT NULL,
                    Kills INT NOT NULL,
                    Deaths INT NOT NULL,
                    Assists INT NOT NULL,
                    HeadshotPercentage FLOAT NOT NULL,
                    Damage INT NOT NULL,
                    CreatedAt DATETIME2 DEFAULT GETUTCDATE(),
                    CONSTRAINT FK_PlayerStats_Demos 
                        FOREIGN KEY (DemoName) REFERENCES Demos(DemoName),
                    CONSTRAINT FK_PlayerStats_Rounds 
                        FOREIGN KEY (DemoName, RoundNumber) 
                        REFERENCES Rounds(DemoName, RoundNumber)
                );
            END;

            -- Create indexes for better query performance
            IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_GrenadeEvents_RoundNumber')
                CREATE INDEX IX_GrenadeEvents_RoundNumber ON GrenadeEvents (DemoName, RoundNumber);

            IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_DeathEvents_RoundNumber')
                CREATE INDEX IX_DeathEvents_RoundNumber ON DeathEvents (DemoName, RoundNumber);

            IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_PlayerStats_RoundNumber_PlayerName')
                CREATE INDEX IX_PlayerStats_RoundNumber_PlayerName ON PlayerStats (DemoName, RoundNumber, PlayerName);";

        await using var command = new SqlCommand(createTablesSql, connection);
        await command.ExecuteNonQueryAsync();
    }
    
    // Helper method to perform bulk insert
    private static async Task BulkInsertData(SqlConnection connection, SqlTransaction transaction,
        string tableName, DataTable dataTable, Dictionary<string, string> columnMappings)
    {
        using var bulkCopy = new SqlBulkCopy(connection, SqlBulkCopyOptions.Default, transaction)
        {
            DestinationTableName = tableName,
            BatchSize = 1000,
            BulkCopyTimeout = 60
        };

        // Set up column mappings
        foreach (var mapping in columnMappings)
        {
            bulkCopy.ColumnMappings.Add(mapping.Key, mapping.Value);
        }

        // Perform the bulk insert
        await bulkCopy.WriteToServerAsync(dataTable);
    }
}