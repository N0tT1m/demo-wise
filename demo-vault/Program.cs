using DemoFile;
using System.Drawing;
using System.Drawing.Imaging;
using ClosedXML.Excel;
using System.Diagnostics;
using DemoFile.Game.Cs;
using DemoFile.Sdk;
using Spectre.Console;
using Color = System.Drawing.Color;

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

internal class Program
{
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
    private static Dictionary<string, List<(float X, float Y, float Z, string Player)>> grenadePositions = new()
    {
        { "Smoke", new List<(float, float, float, string)>() },
        { "Molotov", new List<(float, float, float, string)>() },
        { "HE", new List<(float, float, float, string)>() },
        { "Flashbang", new List<(float, float, float, string)>() }
    };
    
    // Add these new data structures at the class level, alongside the existing grenadePositions dictionary
    private static List<(int roundNum, uint Value, string, string, string Weapon)> grenadeEvents = new();
    private static List<(int roundNum, uint Value, string PlayerName, List<(string GrenadeType, int Count)> inventory)> grenadeInventory = new();
    
    // Update the data structure to store round information
    private static Dictionary<int, RoundInfo> roundTicks = new();

    // Create a class to store round information
    private class RoundInfo
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
    private class PlayerStats
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
    private class DeathEvent
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
    
    private static void CapturePlayerStats(CsDemoParser demo, int roundNum)
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
            Console.WriteLine($"Upper level Z range: {currentMap.DefaultAltitudeMin} to {currentMap.DefaultAltitudeMax}");
            Console.WriteLine($"Lower level Z range: {currentMap.LowerAltitudeMin} to {currentMap.LowerAltitudeMax}");
        }

        var demo = new CsDemoParser();
        var cts = new CancellationTokenSource();
        
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
                StartTick = (int)demo.CurrentGameTick.Value
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
            grenadePositions["Smoke"].Add((e.X, e.Y, e.Z, e.Player?.PlayerName ?? "Unknown"));
            // Console.WriteLine($"{e.Player?.PlayerName} [Smoke] - [{e.X} {e.Y} {e.Z}]; Level: {(e.Z <= currentMap.LowerAltitudeMax ? "Lower" : "Upper")}");
        };
        
        demo.Source1GameEvents.InfernoStartburn += e =>
        {
            // if (e.Player?.PlayerName != "--uncle-") return;
            grenadePositions["Molotov"].Add((e.X, e.Y, e.Z, "Not Available"));
            // Console.WriteLine($"[Molotov] - [{e.X} {e.Y} {e.Z}] Level: {(e.Z <= currentMap.LowerAltitudeMax ? "Lower" : "Upper")}");
        };
        
        demo.Source1GameEvents.HegrenadeDetonate += e =>
        {
            // if (e.Player?.PlayerName != "--uncle-") return;
            grenadePositions["HE"].Add((e.X, e.Y, e.Z, e.Player?.PlayerName ?? "Unknown"));
            // Console.WriteLine($"{e.Player?.PlayerName} [HE] - [{e.X} {e.Y} {e.Z}] Level: {(e.Z <= currentMap.LowerAltitudeMax ? "Lower" : "Upper")}");
        };
        
        demo.Source1GameEvents.FlashbangDetonate += e =>
        {
            // if (e.Player?.PlayerName != "--uncle-") return;
            grenadePositions["Flashbang"].Add((e.X, e.Y, e.Z, e.Player?.PlayerName ?? "Unknown"));
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
            CapturePlayerStats(demo, currentRoundNumber);
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

        demo.EntityEvents.CCSPlayerPawn.AddCollectionChangeCallback(pawn => pawn.Grenades, (pawn, oldGrenades, newGrenades) =>
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
            
            AnsiConsole.MarkupLine($" grenades changed [grey]{string.Join(", ", oldGrenades.Select(x => x.ServerClass.Name))}[/] => [bold]{string.Join(", ", newGrenades.Select(x => x.ServerClass.Name))}[/]");
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
        
                if (player.PlayerPawn is not {} pawn)
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
        
        WriteToCsv(grenadePositions["Flashbang"], "Flashbang.csv", "Flashbang", demo);
        WriteToCsv(grenadePositions["Smoke"], "Smoke.csv", "Smoke", demo);
        WriteToCsv(grenadePositions["HE"], "HE.csv", "HE", demo);

        string outputPath = $"{currentMap.Name}.png";
        string mapImagePath = $"de_{currentMap.Name}.png";

        PlotGrenades(outputPath, mapImagePath);

        if (currentMap.HasLowerLevel)
        {
            outputPath = $"{currentMap.Name}_lower.png";
            mapImagePath = $"de_{currentMap.Name}_lower.png";
            PlotGrenades(outputPath, mapImagePath, true);
        }

        Console.WriteLine($"\nFinished! Check {outputPath} for the plotted grenades.");
    }

    private static void PlotGrenades(string outputPath, string mapImagePath, bool lowerLevel = false)
    {
        try
        {
                Console.WriteLine($"\nProcessing {(lowerLevel ? "lower" : "upper")} level image...");
                
                using var baseImage = new Bitmap(mapImagePath);
                using var bitmap = new Bitmap(baseImage.Width, baseImage.Height, PixelFormat.Format32bppArgb);
                using var graphics = Graphics.FromImage(bitmap);

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
                    foreach (var (x, y, z, player) in positions)
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

                var directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var encoderParameters = new EncoderParameters(1);
                encoderParameters.Param[0] = new EncoderParameter(Encoder.Quality, 100L);
                bitmap.Save(outputPath, GetEncoder(ImageFormat.Png), encoderParameters);
                
                Console.WriteLine($"Saved {outputPath}");
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
    
    private static void ValidateRoundCounts()
    {
        var roundsInStats = playerRoundStats.Select(x => x.Round).Distinct().OrderBy(x => x).ToList();
        var roundsInTicks = roundTicks.Keys.OrderBy(x => x).ToList();
        var roundsInEvents = grenadeEvents.Select(x => x.roundNum).Distinct().OrderBy(x => x).ToList();
    
        Console.WriteLine("\nValidating round counts:");
        Console.WriteLine($"Rounds in player stats: {roundsInStats.Count} ({string.Join(", ", roundsInStats)})");
        Console.WriteLine($"Rounds in ticks: {roundsInTicks.Count} ({string.Join(", ", roundsInTicks)})");
        Console.WriteLine($"Rounds in events: {roundsInEvents.Count} ({string.Join(", ", roundsInEvents)})");
    
        if (roundsInStats.Count != roundsInTicks.Count)
        {
            Console.WriteLine("WARNING: Mismatch in round counts!");
            var missingRounds = roundsInTicks.Except(roundsInStats).ToList();
            if (missingRounds.Any())
            {
                Console.WriteLine($"Missing rounds in stats: {string.Join(", ", missingRounds)}");
            }
        }
    }
    
    // 3. Update the WritePlayerStatsToExcel method with debug logging
    private static void WritePlayerStatsToExcel(XLWorkbook workbook)
    {
        // Add this call just before writing to Excel:
        ValidateRoundCounts();
        
        // Clean up round numbers before writing
        CleanupRoundNumbers();
        
        Console.WriteLine($"\nWriting player stats to Excel. Total stats entries: {playerRoundStats.Count}");
        var statsSheet = workbook.Worksheets.Add("Player Stats");
        
        // Write headers
        statsSheet.Cell(1, 1).Value = "Round";
        statsSheet.Cell(1, 2).Value = "Player Name";
        statsSheet.Cell(1, 3).Value = "Steam ID";
        statsSheet.Cell(1, 4).Value = "Team";
        statsSheet.Cell(1, 5).Value = "Ping";
        statsSheet.Cell(1, 6).Value = "Money";
        statsSheet.Cell(1, 7).Value = "Kills";
        statsSheet.Cell(1, 8).Value = "Deaths";
        statsSheet.Cell(1, 9).Value = "Assists";
        statsSheet.Cell(1, 10).Value = "Headshot %";
        statsSheet.Cell(1, 11).Value = "Damage";

        // Sort by round number then player name for consistent ordering
        var sortedStats = playerRoundStats.OrderBy(x => x.Round)
                                        .ThenBy(x => x.PlayerName)
                                        .ToList();

        var rowIndex = 2;
        foreach (var stat in sortedStats)
        {
            Console.WriteLine($"Writing row {rowIndex}: Round {stat.Round} - {stat.PlayerName}");
        
            statsSheet.Cell(rowIndex, 1).Value = stat.Round;
            statsSheet.Cell(rowIndex, 2).Value = stat.PlayerName;
            statsSheet.Cell(rowIndex, 3).Value = stat.SteamID;
            statsSheet.Cell(rowIndex, 4).Value = stat.Team;
            statsSheet.Cell(rowIndex, 5).Value = stat.Ping;
            statsSheet.Cell(rowIndex, 6).Value = stat.Money;
            statsSheet.Cell(rowIndex, 7).Value = stat.Kills;
            statsSheet.Cell(rowIndex, 8).Value = stat.Deaths;
            statsSheet.Cell(rowIndex, 9).Value = stat.Assists;
            statsSheet.Cell(rowIndex, 10).Value = stat.HeadshotPercentage;
            statsSheet.Cell(rowIndex, 11).Value = stat.Damage;
            
            rowIndex++;
        }
        
        Console.WriteLine($"Finished writing {rowIndex - 2} rows of player stats");
        statsSheet.Columns().AdjustToContents();
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
        // Get the min and max round numbers from player stats to handle dynamic round count
        var validRounds = playerRoundStats.Select(x => x.Round).Distinct().OrderBy(x => x).ToList();
        var minRound = validRounds.Min();
        var maxRound = validRounds.Max();
        
        Console.WriteLine($"\nFound rounds from {minRound} to {maxRound}");
        
        var roundMapping = new Dictionary<int, int>();
        int newRoundNumber = 1; // Start at 1
        
        // Create mapping for each round we've seen
        foreach (var round in validRounds)
        {
            roundMapping[round] = newRoundNumber++;
            Console.WriteLine($"Mapping round {round} -> {roundMapping[round]}");
        }

        // Update grenade events
        grenadeEvents = grenadeEvents
            .Where(e => roundMapping.ContainsKey(e.roundNum))
            .Select(e => (roundMapping[e.roundNum], e.Value, e.Item3, e.Item4, e.Weapon))
            .ToList();

        Console.WriteLine($"Updated {grenadeEvents.Count} grenade events");

        // Update grenade inventory
        grenadeInventory = grenadeInventory
            .Where(inv => roundMapping.ContainsKey(inv.roundNum))
            .Select(inv => (roundMapping[inv.roundNum], inv.Value, inv.PlayerName, inv.inventory))
            .ToList();

        Console.WriteLine($"Updated {grenadeInventory.Count} inventory records");

        // Update death events
        deathEvents = deathEvents
            .Where(d => roundMapping.ContainsKey(d.RoundNumber))
            .Select(d => new DeathEvent
            {
                RoundNumber = roundMapping[d.RoundNumber],
                Tick = d.Tick,
                AttackerName = d.AttackerName,
                AttackerTeam = d.AttackerTeam,
                VictimName = d.VictimName,
                VictimTeam = d.VictimTeam,
                Weapon = d.Weapon,
                Headshot = d.Headshot
            })
            .ToList();

        Console.WriteLine($"Updated {deathEvents.Count} death events");

        // Update round ticks
        var updatedRoundTicks = new Dictionary<int, RoundInfo>();
        foreach (var (round, info) in roundTicks.Where(kvp => roundMapping.ContainsKey(kvp.Key)))
        {
            updatedRoundTicks[roundMapping[round]] = info;
        }
        roundTicks = updatedRoundTicks;

        Console.WriteLine($"Updated {roundTicks.Count} round records");

        // Print final round distribution
        var finalRoundCounts = new Dictionary<string, List<int>>
        {
            { "Grenade Events", grenadeEvents.Select(x => x.roundNum).Distinct().OrderBy(x => x).ToList() },
            { "Inventory", grenadeInventory.Select(x => x.roundNum).Distinct().OrderBy(x => x).ToList() },
            { "Deaths", deathEvents.Select(x => x.RoundNumber).Distinct().OrderBy(x => x).ToList() },
            { "Round Ticks", roundTicks.Keys.OrderBy(x => x).ToList() }
        };

        Console.WriteLine("\nFinal round distribution:");
        foreach (var (category, rounds) in finalRoundCounts)
        {
            Console.WriteLine($"{category}: {string.Join(", ", rounds)}");
        }
    }
    
    private static void WriteToCsv(List<(float X, float Y, float Z, string Player)> lines, string outputPath, string type, CsDemoParser demo)
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
                foreach (var (x, y, z, player) in grenadePositions[grenadeType])
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
}