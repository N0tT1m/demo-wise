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
    // Add at the class level
    private static Dictionary<int, (int StartTick, int FreezeEndTick, int EndTick)> roundTicks = new();


    private static MapConfig currentMap;

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

        demo.Source1GameEvents.SmokegrenadeDetonate += e =>
        {
            // if (e.Player?.PlayerName != "--uncle-") return;
            grenadePositions["Smoke"].Add((e.X, e.Y, e.Z, e.Player?.PlayerName ?? "Unknown"));
            Console.WriteLine($"{e.Player?.PlayerName} [Smoke] - [{e.X} {e.Y} {e.Z}]; Level: {(e.Z <= currentMap.LowerAltitudeMax ? "Lower" : "Upper")}");
        };
        
        demo.Source1GameEvents.MolotovDetonate += e =>
        {
            // if (e.Player?.PlayerName != "--uncle-") return;
            grenadePositions["Molotov"].Add((e.X, e.Y, e.Z, e.Player?.PlayerName ?? "Unknown"));
            Console.WriteLine($"{e.Player?.PlayerName} [Molotov] - [{e.X} {e.Y} {e.Z}] Level: {(e.Z <= currentMap.LowerAltitudeMax ? "Lower" : "Upper")}");
        };
        
        demo.Source1GameEvents.HegrenadeDetonate += e =>
        {
            // if (e.Player?.PlayerName != "--uncle-") return;
            grenadePositions["HE"].Add((e.X, e.Y, e.Z, e.Player?.PlayerName ?? "Unknown"));
            Console.WriteLine($"{e.Player?.PlayerName} [HE] - [{e.X} {e.Y} {e.Z}] Level: {(e.Z <= currentMap.LowerAltitudeMax ? "Lower" : "Upper")}");
        };
        
        demo.Source1GameEvents.FlashbangDetonate += e =>
        {
            // if (e.Player?.PlayerName != "--uncle-") return;
            grenadePositions["Flashbang"].Add((e.X, e.Y, e.Z, e.Player?.PlayerName ?? "Unknown"));
            Console.WriteLine($"{e.Player?.PlayerName} [Flashbang] - [{e.X} {e.Y} {e.Z}] Level: {(e.Z <= currentMap.LowerAltitudeMax ? "Lower" : "Upper")}");
        };

        var roundNum = 0;
        // Update the round event handlers
        demo.Source1GameEvents.RoundStart += e =>
        {
            roundNum += 1;
            Console.WriteLine($"\n\n>>> Round start [{roundNum}] <<<");
    
            // If previous round exists and doesn't have an end tick, set it
            if (roundNum > 1 && roundTicks.ContainsKey(roundNum - 1))
            {
                var prevRound = roundTicks[roundNum - 1];
                if (prevRound.EndTick == 0)
                {
                    prevRound.EndTick = (int)demo.CurrentGameTick.Value;
                    roundTicks[roundNum - 1] = prevRound;
                }
            }
    
            roundTicks[roundNum] = ((int StartTick, int FreezeEndTick, int EndTick))(demo.CurrentGameTick.Value, 0, 0);
        };

        demo.Source1GameEvents.RoundFreezeEnd += e =>
        {
            Console.WriteLine("\n  > Round freeze end");
            if (roundTicks.ContainsKey(roundNum))
            {
                var record = roundTicks[roundNum];
                record.FreezeEndTick = (int)demo.CurrentGameTick.Value;
                roundTicks[roundNum] = record;
            }
            DumpGrenadeInventory();
        };
        
        demo.Source1GameEvents.RoundEnd += e =>
        {
            Console.WriteLine("\n  > Round end");
            if (roundTicks.ContainsKey(roundNum))
            {
                var record = roundTicks[roundNum];
                record.EndTick = (int)demo.CurrentGameTick.Value;
                roundTicks[roundNum] = record;
            }
            DumpGrenadeInventory();
        };

        demo.EntityEvents.CCSPlayerPawn.AddCollectionChangeCallback(pawn => pawn.Grenades, (pawn, oldGrenades, newGrenades) =>
        {
            Console.Write($"  [Tick {demo.CurrentGameTick.Value}] ");
            MarkupPlayerName(pawn.Controller);
            var playerName = pawn.Controller?.PlayerName ?? "Unknown";
            
            // Store the grenade changes
            grenadeEvents.Add((
                roundNum,
                demo.CurrentGameTick.Value,
                playerName,
                "InventoryChange",
                string.Join(", ", newGrenades.Select(x => x.ServerClass.Name))
            ));
            
            AnsiConsole.MarkupLine($" grenades changed [grey]{string.Join(", ", oldGrenades.Select(x => x.ServerClass.Name))}[/] => [bold]{string.Join(", ", newGrenades.Select(x => x.ServerClass.Name))}[/]");
        });

        demo.Source1GameEvents.WeaponFire += e =>
        {
            if (!e.Weapon.Contains("nade") && !e.Weapon.Contains("molotov"))
                return;

            Console.Write($"  [Tick {demo.CurrentGameTick.Value}] ");
            MarkupPlayerName(e.Player);
            
            // Store the throw event
            grenadeEvents.Add((
                roundNum,
                demo.CurrentGameTick.Value,
                e.Player?.PlayerName ?? "Unknown",
                "Throw",
                e.Weapon
            ));
            
            AnsiConsole.MarkupLine($" [bold]threw a {e.Weapon}[/]");
        };

        void DumpGrenadeInventory()
        {
            // Store inventory for all players, even if they have no grenades
            foreach (var player in demo.Players)
            {
                var inventory = new List<(string GrenadeType, int Count)>();
                Console.Write("    ");
                MarkupPlayerName(player);
                Console.Write(" - ");

                if (player.PlayerPawn is not {} pawn)
                {
                    if (player.PlayerName == "SourceTV")
                    {
                        continue;
                    } 
                    
                    Console.WriteLine("<no pawn>");
                    inventory.Add(("NO_PAWN", 0));
                }
                else if (!pawn.IsAlive)
                {
                    Console.WriteLine("<dead>");
                    inventory.Add(("DEAD", 0));
                }
                else
                {
                    var grenades = pawn.Grenades;
                    if (!grenades.Any())
                    {
                        Console.WriteLine("<no grenades>");
                        inventory.Add(("NO_GRENADES", 0));
                    }
                    else
                    {
                        foreach (var grenade in grenades)
                        {
                            inventory.Add((grenade.ServerClass.Name, grenade.GrenadeCount));
                            Console.Write($"{grenade.ServerClass.Name} x {grenade.GrenadeCount}, ");
                        }
                        Console.WriteLine("");
                    }
                }

                // Store the inventory for every player in every round
                grenadeInventory.Add((
                    roundNum,
                    demo.CurrentGameTick.Value,
                    player.PlayerName,
                    inventory
                ));
            }
        }

        var ticks = demo.CurrentDemoTick.Value;

        var reader = DemoFileReader.Create(demo, File.OpenRead(demoPath));
        await reader.ReadAllAsync();
        
        WriteToCsv(grenadePositions["Flashbang"], "Flashbang.csv", "Flashbang");
        WriteToCsv(grenadePositions["Smoke"], "Smoke.csv", "Smoke");
        WriteToCsv(grenadePositions["HE"], "HE.csv", "HE");

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

    // Update WriteToCsv to include the new data
    private static void WriteToCsv(List<(float X, float Y, float Z, string Player)> lines, string outputPath, string type)
    {
        try
        {
            // Store data until all types are processed
            if (type != "Flashbang") // Use Flashbang as the trigger since it's the last one processed
            {
                return;
            }

            string excelPath = "grenades.xlsx";
            using var workbook = new XLWorkbook();

            // Original grenade positions sheets
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

            // Add Grenade Events sheet
            var eventsSheet = workbook.Worksheets.Add("Grenade Events");
            eventsSheet.Cell(1, 1).Value = "Round";
            eventsSheet.Cell(1, 2).Value = "Tick";
            eventsSheet.Cell(1, 3).Value = "Player";
            eventsSheet.Cell(1, 4).Value = "Action";
            eventsSheet.Cell(1, 5).Value = "Grenade Type";

            var eventRowIndex = 2;
            foreach (var (round, tick, player, action, grenadeType) in grenadeEvents)
            {
                eventsSheet.Cell(eventRowIndex, 1).Value = round;
                eventsSheet.Cell(eventRowIndex, 2).Value = tick;
                eventsSheet.Cell(eventRowIndex, 3).Value = player;
                eventsSheet.Cell(eventRowIndex, 4).Value = action;
                eventsSheet.Cell(eventRowIndex, 5).Value = grenadeType;
                eventRowIndex++;
            }
            eventsSheet.Columns().AdjustToContents();

            // Add Inventory Snapshots sheet
            var inventorySheet = workbook.Worksheets.Add("Inventory Snapshots");
            inventorySheet.Cell(1, 1).Value = "Round";
            inventorySheet.Cell(1, 2).Value = "Tick";
            inventorySheet.Cell(1, 3).Value = "Player";
            inventorySheet.Cell(1, 4).Value = "Inventory";

            var invRowIndex = 2;
            foreach (var (round, tick, player, inventory) in grenadeInventory)
            {
                inventorySheet.Cell(invRowIndex, 1).Value = round;
                inventorySheet.Cell(invRowIndex, 2).Value = tick;
                inventorySheet.Cell(invRowIndex, 3).Value = player;
                inventorySheet.Cell(invRowIndex, 4).Value = string.Join(", ", inventory.Select(x => $"{x.GrenadeType} x {x.Count}"));
                invRowIndex++;
            }
            inventorySheet.Columns().AdjustToContents();

            // Update the Excel writing portion for the Rounds sheet
            var roundsSheet = workbook.Worksheets.Add("Rounds");
            roundsSheet.Cell(1, 1).Value = "Round";
            roundsSheet.Cell(1, 2).Value = "Start Tick";
            roundsSheet.Cell(1, 3).Value = "Freeze End Tick";
            roundsSheet.Cell(1, 4).Value = "End Tick";
            roundsSheet.Cell(1, 5).Value = "Round Duration";
            roundsSheet.Cell(1, 6).Value = "Freeze Time";

            var roundRowIndex = 2;
            foreach (var (round, (startTick, freezeEndTick, endTick)) in roundTicks.OrderBy(x => x.Key))
            {
                if (endTick > 0) // Only include complete rounds
                {
                    roundsSheet.Cell(roundRowIndex, 1).Value = round;
                    roundsSheet.Cell(roundRowIndex, 2).Value = startTick;
                    roundsSheet.Cell(roundRowIndex, 3).Value = freezeEndTick;
                    roundsSheet.Cell(roundRowIndex, 4).Value = endTick;
                    roundsSheet.Cell(roundRowIndex, 5).Value = endTick - startTick;
                    roundsSheet.Cell(roundRowIndex, 6).Value = freezeEndTick > 0 ? freezeEndTick - startTick : 0;
                    roundRowIndex++;
                }
            }
            roundsSheet.Columns().AdjustToContents();

            // Create directory if it doesn't exist
            var directory = Path.GetDirectoryName(excelPath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook
            workbook.SaveAs(excelPath);
            Console.WriteLine($"\nSuccessfully wrote all data to {excelPath}");
            
            // Print summary
            Console.WriteLine("\nSummary:");
            foreach (var (grenadeType, positions) in grenadePositions)
            {
                Console.WriteLine($"- {grenadeType} positions: {positions.Count}");
            }
            Console.WriteLine($"- Grenade events: {grenadeEvents.Count}");
            Console.WriteLine($"- Inventory snapshots: {grenadeInventory.Count}");
            Console.WriteLine($"- Rounds recorded: {roundTicks.Count}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error writing to Excel file: {ex.Message}");
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