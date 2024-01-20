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
        private static string modelLocation = model.GetInfo().ModelPath;
        public static List<string> profiles = new List<string>() { "IPE80", "IPE100", "IPE120", "IPE140", "IPE160", "IPE180", "IPE200", "IPE220", "IPE240", "IPE270", "IPE300", "IPE330", "IPE360", "IPE400", "IPE450", "IPE500", "IPE550", "IPE600" };

        static void Main(string[] args)
        {
            if (!model.GetConnectionStatus())
                return;

            //GenerateExampleBeams();          

            ColorByProfile();
        }

        private static void ColorByProfile()
        {
            foreach (string profile in profiles)
            {
                CreateObjectGroupFile(profile);
            }

            CreateRepresentationFile();
        }

        private static void CreateRepresentationFile()
        {
            string filePath = Path.Combine(modelLocation, "attributes", "#ColorByProfile.rep");
            int itemCount = profiles.Count;
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
            stringBuilder.AppendLine("SECTION_OBJECT_REP");
            stringBuilder.AppendLine("{");
            stringBuilder.AppendLine($"OG_{profiles[index].Replace('/', '_')}");
            stringBuilder.AppendLine($"{index + 13969144}");
            stringBuilder.AppendLine("10");
            stringBuilder.AppendLine("}");
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
            (SETTINGNOTDEFINED) 
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
            string objectGroupFilePath = Path.Combine(modelLocation, "attributes", "OG_" + profile.Replace('/', '_') + ".PObjGrp");
            string content =   $"TITLE_OBJECT_GROUP\n" +
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

        private static int[][] GenerateRainbowColorsWithoutRed(int n)
        {
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

        private static int[] HsvToRgb(double hue, double saturation, double value)
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

        private static void GenerateExampleBeams()
        {
            Point A = new Point(0, 0, 0);
            Point B = new Point(2000, 0, 0);

            foreach(string profile in profiles)
            {
                Beam beam = new Beam();
                beam.Profile.ProfileString = profile;
                beam.StartPoint = A;
                beam.EndPoint = B;
                A.Y += 500;
                B.Y += 500;
                beam.Insert();
            }
            model.CommitChanges();
        }

        private static string CreateObjectGroupFile(string proWHAT, string alblWhat, string value)
        {
            string result = $"TITLE_OBJECT_GROUP\n" +
                            "{\n" +
                            "    Version= 1.05\n" +
                            "    Count= 1\n" +
                            "    SECTION_OBJECT_GROUP\n" +
                            "    {\n" +
                            "        0\n" +
                            "        1\n" +
                            "        co_part\n" +
                            $"        {proWHAT}\n" +
                            $"        {alblWhat}\n" +
                            "        ==\n" +
                            "        albl_Equals\n" +
                            $"        {value}\n" +
                            "        0\n" +
                            "        &&\n" +
                            "    }\n" +
                            "}";

            return result;
        }
    
    }
}
