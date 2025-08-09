using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;

namespace ColorByProfile
{
    internal class Program
    {
        private static Model model = new Model();
        private static string modelLocation = string.Empty;
        public static List<string> profiles = new List<string>();

        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("=== ColorByProfile - Start ===");

                // Sprawdź połączenie z modelem
                if (!model.GetConnectionStatus())
                {
                    Console.WriteLine("BŁĄD: Brak połączenia z modelem Tekla Structures.");
                    Console.WriteLine("Upewnij się, że Tekla Structures jest uruchomiona i model jest otwarty.");
                    Console.WriteLine("Naciśnij dowolny klawisz aby zakończyć...");
                    Console.ReadKey();
                    return;
                }

                Console.WriteLine("✓ Połączenie z modelem nawiązane pomyślnie.");

                // Pobierz lokalizację modelu
                try
                {
                    modelLocation = model.GetInfo().ModelPath;
                    Console.WriteLine($"✓ Lokalizacja modelu: {modelLocation}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"BŁĄD: Nie można pobrać ścieżki modelu: {ex.Message}");
                    Console.WriteLine("Naciśnij dowolny klawisz aby zakończyć...");
                    Console.ReadKey();
                    return;
                }

                // Pobierz profile z modelu
                Console.WriteLine("Pobieranie profili z modelu...");
                profiles = GetUsedProfilesFromModel();

                if (profiles.Count == 0)
                {
                    Console.WriteLine("OSTRZEŻENIE: Nie znaleziono żadnych profili w modelu.");
                    Console.WriteLine("Sprawdź czy model zawiera elementy konstrukcyjne (belki, słupy).");
                    Console.WriteLine("Naciśnij dowolny klawisz aby zakończyć...");
                    Console.ReadKey();
                    return;
                }

                Console.WriteLine($"✓ Znaleziono {profiles.Count} różnych profili w modelu:");
                foreach (string profile in profiles)
                {
                    Console.WriteLine($"  - {profile}");
                }

                // Główna funkcjonalność
                Console.WriteLine("\nTworzenie kolorowania według profili...");
                ColorByProfile();

                Console.WriteLine("✓ Kolorowanie według profili zostało utworzone pomyślnie!");
                Console.WriteLine("=== ColorByProfile - Koniec ===");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"KRYTYCZNY BŁĄD: {ex.Message}");
                Console.WriteLine($"Szczegóły: {ex.StackTrace}");
            }
            finally
            {
                Console.WriteLine("\nNaciśnij dowolny klawisz aby zakończyć...");
                Console.ReadKey();
            }
        }

        private static void ColorByProfile()
        {
            try
            {
                Console.WriteLine("Tworzenie plików grup obiektów...");
                int createdFiles = 0;

                foreach (string profile in profiles)
                {
                    try
                    {
                        CreateObjectGroupFile(profile);
                        createdFiles++;
                        Console.WriteLine($"  ✓ Utworzono grupę dla profilu: {profile}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  BŁĄD przy tworzeniu grupy dla profilu '{profile}': {ex.Message}");
                    }
                }

                Console.WriteLine($"✓ Utworzono {createdFiles} plików grup obiektów.");

                Console.WriteLine("Tworzenie pliku reprezentacji...");
                CreateRepresentationFile();
                Console.WriteLine("✓ Plik reprezentacji utworzony pomyślnie.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"BŁĄD w ColorByProfile(): {ex.Message}");
                throw;
            }
        }

        private static List<string> GetUsedProfilesFromModel()
        {
            HashSet<string> uniqueProfiles = new HashSet<string>();
            int processedObjects = 0;
            int profileObjects = 0;

            try
            {
                ModelObjectEnumerator modelObjectEnum = model.GetModelObjectSelector().GetAllObjects();

                while (modelObjectEnum.MoveNext())
                {
                    try
                    {
                        ModelObject obj = modelObjectEnum.Current;
                        processedObjects++;

                        if (obj is Beam beam)
                        {
                            string profileString = beam.Profile.ProfileString;
                            if (profileString.Contains('*'))
                                profileString = profileString.Split('*')[0] + "*";
                            if (profileString.Contains(' '))
                                profileString = profileString.Split(' ')[0] + "*";
                            if (profileString.StartsWith("BL") || profileString.StartsWith("PL"))
                                profileString = profileString.Split('*')[0] + "*";

                            if (!string.IsNullOrEmpty(profileString))
                            {
                                uniqueProfiles.Add(profileString);
                                profileObjects++;
                            }
                        }
                        else if (obj is ContourPlate plate)
                        {
                            string profileString = plate.Profile.ProfileString;
                            if (profileString.Contains('*'))
                                profileString = profileString.Split('*')[0] + "*";
                            if (profileString.Contains(' '))
                                profileString = profileString.Split(' ')[0] + "*";
                            if (profileString.StartsWith("BL") || profileString.StartsWith("PL"))
                                profileString = profileString.Split('*')[0] + "*";

                            if (!string.IsNullOrEmpty(profileString))
                            {
                                uniqueProfiles.Add(profileString);
                                profileObjects++;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"OSTRZEŻENIE: Błąd przy przetwarzaniu obiektu: {ex.Message}");
                    }
                }

                Console.WriteLine($"✓ Przetworzono {processedObjects} obiektów, znaleziono {profileObjects} elementów z profilami.");
                Console.WriteLine($"✓ Znaleziono {uniqueProfiles.Count} unikalnych profili.");

                // Sortowanie alfabetyczne
                var sortedProfiles = uniqueProfiles.OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToList();
                Console.WriteLine("✓ Profile posortowane alfabetycznie.");

                return sortedProfiles;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"BŁĄD przy pobieraniu profili z modelu: {ex.Message}");
                throw;
            }
        }

        private static void CreateRepresentationFile()
        {
            try
            {
                string attributesPath = Path.Combine(modelLocation, "attributes");

                // Sprawdź czy folder attributes istnieje
                if (!Directory.Exists(attributesPath))
                {
                    Console.WriteLine($"Tworzenie folderu: {attributesPath}");
                    Directory.CreateDirectory(attributesPath);
                }

                string filePath = Path.Combine(attributesPath, "1_profile.rep");
                int itemCount = profiles.Count;

                Console.WriteLine($"Generowanie {itemCount} kolorów...");
                int[][] colors = GenerateRainbowColorsWithoutRed(itemCount);

                StringBuilder stringBuilder = new StringBuilder();

                stringBuilder.AppendLine($"REPRESENTATIONS");
                stringBuilder.AppendLine("{");
                stringBuilder.AppendLine("Version= 1.04");
                stringBuilder.AppendLine($"Count= {itemCount + 1}");

                for (int i = 0; i < itemCount; i++)
                {
                    AppendSectionUtilityLimits(stringBuilder);
                    AppendSectionObjectRep(stringBuilder, i);
                    AppendSectionObjectRepByAttribute(stringBuilder);
                    AppendSectionObjectRepRgbValue(stringBuilder, colors[i]);
                }

                AppendDefaultSection(stringBuilder);

                File.WriteAllText(filePath, stringBuilder.ToString());
                Console.WriteLine($"✓ Plik reprezentacji zapisany: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"BŁĄD przy tworzeniu pliku reprezentacji: {ex.Message}");
                throw;
            }
        }

        private static void AppendSectionUtilityLimits(StringBuilder stringBuilder)
        {
            stringBuilder.AppendLine("SECTION_UTILITY_LIMITS");
            stringBuilder.AppendLine("{");
            stringBuilder.AppendLine("0.5");
            stringBuilder.AppendLine("0.9");
            stringBuilder.AppendLine("1");
            stringBuilder.AppendLine("1.2");
            stringBuilder.AppendLine("}");
        }

        private static void AppendSectionObjectRep(StringBuilder stringBuilder, int index)
        {
            try
            {
                stringBuilder.AppendLine("SECTION_OBJECT_REP");
                stringBuilder.AppendLine("{");
                stringBuilder.AppendLine($"OG_{profiles[index]}");
                stringBuilder.AppendLine($"{index + 13969144}");
                stringBuilder.AppendLine("10");
                stringBuilder.AppendLine("}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"BŁĄD przy tworzeniu sekcji reprezentacji dla indeksu {index}: {ex.Message}");
                throw;
            }
        }

        private static void AppendSectionObjectRepByAttribute(StringBuilder stringBuilder)
        {
            stringBuilder.AppendLine("SECTION_OBJECT_REP_BY_ATTRIBUTE");
            stringBuilder.AppendLine("{");
            stringBuilder.AppendLine("(SETTING\u0002NOT\u0002DEFINED)");
            stringBuilder.AppendLine("}");
        }

        private static void AppendSectionObjectRepRgbValue(StringBuilder stringBuilder, int[] color)
        {
            stringBuilder.AppendLine("SECTION_OBJECT_REP_RGB_VALUE");
            stringBuilder.AppendLine("{");
            stringBuilder.AppendLine(color[0].ToString());
            stringBuilder.AppendLine(color[1].ToString());
            stringBuilder.AppendLine(color[2].ToString());
            stringBuilder.AppendLine("}");
        }

        private static void AppendDefaultSection(StringBuilder stringBuilder)
        {
            stringBuilder.AppendLine(@"SECTION_UTILITY_LIMITS 
        {
            0 
            0 
            0 
            0 
        }
        SECTION_OBJECT_REP 
        {
            All 
            1 
            10 
        }
        SECTION_OBJECT_REP_BY_ATTRIBUTE 
        {
            (SETTINGNOTDEFINED) 
        }
        SECTION_OBJECT_REP_RGB_VALUE 
        {
            -1 
            -1 
            -1 
        }
        }
        ");
        }

        private static void CreateObjectGroupFile(string profile)
        {
            try
            {
                string attributesPath = Path.Combine(modelLocation, "attributes");

                // Sprawdź czy folder attributes istnieje
                if (!Directory.Exists(attributesPath))
                {
                    Directory.CreateDirectory(attributesPath);
                }

                string objectGroupFilePath = Path.Combine(attributesPath, "OG_" + profile.Replace('/', '_').Replace('*', '_').Replace(' ', '_') + ".PObjGrp");
                string content = $"TITLE_OBJECT_GROUP\n" +
                                "{\n" +
                                "    Version= 1.05\n" +
                                "    Count= 1\n" +
                                "    SECTION_OBJECT_GROUP\n" +
                                "    {\n" +
                                "        0\n" +
                                "        1\n" +
                                "        co_part\n" +
                                $"        proPROFILE\n" +
                                $"        albl_Profile\n" +
                                "        ==\n" +
                                "        albl_Equals\n" +
                                $"        {profile}\n" +
                                "        0\n" +
                                "        &&\n" +
                                "    }\n" +
                                "}";

                File.WriteAllText(objectGroupFilePath, content);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"BŁĄD przy tworzeniu pliku grupy dla profilu '{profile}': {ex.Message}");
                throw;
            }
        }

        private static int[][] GenerateRainbowColorsWithoutRed(int n)
        {
            try
            {
                if (n <= 0)
                {
                    Console.WriteLine("OSTRZEŻENIE: Liczba kolorów musi być większa od 0");
                    return new int[0][];
                }

                int[][] colors = new int[n][];

                double hueStep = 300.0 / n;
                double startHue = 60.0;

                for (int i = 0; i < n; i++)
                {
                    double hue = startHue + (i * hueStep);
                    int[] color = HsvToRgb(hue, 1.0, 1.0);
                    colors[i] = color;
                }
                return colors;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"BŁĄD przy generowaniu kolorów: {ex.Message}");
                throw;
            }
        }

        private static int[] HsvToRgb(double hue, double saturation, double value)
        {
            try
            {
                double c = value * saturation;
                double x = c * (1 - Math.Abs((hue / 60) % 2 - 1));
                double m = value - c;

                double red, green, blue;

                if (hue >= 60 && hue < 120)
                {
                    red = x;
                    green = c;
                    blue = 0;
                }
                else if (hue >= 120 && hue < 180)
                {
                    red = 0;
                    green = c;
                    blue = x;
                }
                else if (hue >= 180 && hue < 240)
                {
                    red = 0;
                    green = x;
                    blue = c;
                }
                else if (hue >= 240 && hue < 300)
                {
                    red = x;
                    green = 0;
                    blue = c;
                }
                else
                {
                    red = c;
                    green = 0;
                    blue = x;
                }

                int[] rgb = new int[]
                {
                (int)((red + m) * 255),
                (int)((green + m) * 255),
                (int)((blue + m) * 255)
                };

                return rgb;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"BŁĄD przy konwersji HSV do RGB: {ex.Message}");
                throw;
            }
        }

        private static void GenerateExampleBeams()
        {
            try
            {
                Console.WriteLine("Generowanie przykładowych belek...");
                Point A = new Point(0, 0, 0);
                Point B = new Point(2000, 0, 0);

                foreach (string profile in profiles)
                {
                    try
                    {
                        Beam beam = new Beam();
                        beam.Profile.ProfileString = profile;
                        beam.StartPoint = A;
                        beam.EndPoint = B;
                        A.Y += 500;
                        B.Y += 500;
                        beam.Insert();
                        Console.WriteLine($"✓ Utworzono belkę z profilem: {profile}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"BŁĄD przy tworzeniu belki z profilem '{profile}': {ex.Message}");
                    }
                }
                model.CommitChanges();
                Console.WriteLine("✓ Zmiany w modelu zostały zatwierdzone.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"BŁĄD przy generowaniu przykładowych belek: {ex.Message}");
                throw;
            }
        }
    }
}