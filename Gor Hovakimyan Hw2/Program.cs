using OfficeOpenXml;
using System.Drawing;

namespace HW2
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string folderPath = @"C:\Users\gorho\Desktop\Rafik hm2\face";
            string excelPath = @"C:\Users\gorho\Desktop\Rafik hm2\RGB Rafik 6.xlsx";

            var colorsSet = new SortedSet<int>();

            // Read images and collect unique non-black face-specific colors
            foreach (var filePath in Directory.GetFiles(folderPath, "*.png")) // or "*.jpg" depending on the image format
            {
                using (var image = new Bitmap(filePath))
                {
                    for (int y = 0; y < image.Height; y++)
                    {
                        for (int x = 0; x < image.Width; x++)
                        {
                            var color = image.GetPixel(x, y);
                            if (color.R != 0 || color.G != 0 || color.B != 0) // Non-black check
                            {
                                int colorValue = color.ToArgb();
                                colorsSet.Add(colorValue);
                            }
                        }
                    }
                }
            }

            // Initialize arrays for statistics
            int[] count = new int[256];
            double[] min = new double[256];
            double[] max = new double[256];
            double[] mean = new double[256];
            double[] mean2 = new double[256];

            for (int i = 0; i < 256; i++)
            {
                min[i] = double.MaxValue;
                max[i] = double.MinValue;
            }

            // Compute statistics
            foreach (var colorValue in colorsSet)
            {
                Color color = Color.FromArgb(colorValue);
                int G = color.G;
                double rbAvg = (color.R + color.B) / 2.0;

                count[G]++;
                if (rbAvg < min[G]) min[G] = rbAvg;
                if (rbAvg > max[G]) max[G] = rbAvg;
                mean[G] += rbAvg;
                mean2[G] += rbAvg * rbAvg;
            }

            for (int i = 0; i < 256; i++)
            {
                if (count[i] > 0)
                {
                    mean[i] /= count[i];
                    mean2[i] /= count[i];
                }
                else
                {
                    min[i] = 0;
                    max[i] = 0;
                }
            }

            // Write results to Excel
            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                var worksheet = package.Workbook.Worksheets["RB(G) 11"];
                for (int i = 1; i < 256; i++)
                {
                    worksheet.Cells[i + 2, 3].Value = count[i];
                    worksheet.Cells[i + 2, 4].Value = min[i];
                    worksheet.Cells[i + 2, 5].Value = max[i];
                    worksheet.Cells[i + 2, 6].Value = mean[i];
                    worksheet.Cells[i + 2, 7].Value = mean2[i];
                }
                package.Save();
            }

            Console.WriteLine("Excel file updated successfully.");
        }
    }
}
