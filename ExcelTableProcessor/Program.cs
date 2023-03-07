using System;
using System.IO;
using System.Windows.Forms;
using ExcelDataReader;

class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        // Prompt the user to select a file
        var openFileDialog = new OpenFileDialog();
        openFileDialog.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx";
        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            var excelFilePath = openFileDialog.FileName;
            var desktopFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var outputFilePath = Path.Combine(desktopFolderPath, "Output.txt");


            using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                // Read the header row
                reader.Read();

                // Open the output file for writing
                using (var writer = new StreamWriter(outputFilePath))
                {
                    while (reader.Read())
                    {
                        var codeObj = reader.GetValue(0);
                        var dayObj = reader.GetValue(1);
                        var slotObj = reader.GetValue(2);

                        if (codeObj != null && dayObj != null && slotObj != null)
                        {
                            var code = codeObj.ToString().Trim();
                            var day = dayObj.ToString().Trim();
                            var slot = slotObj.ToString().Trim();

                            var firstCharDay = day.Substring(0, 1);
                            var firstCharSlot = slot.Substring(0, 1);

                            // Write the row with the first character of day concatenated with the slot
                            writer.WriteLine($"{code}\t{firstCharDay}{firstCharSlot}");

                            // Write additional rows if the day and slot have more than one character
                            for (int i = 1; i < Math.Min(day.Length, slot.Length); i++)
                            {
                                var charDay = day.Substring(i, 1);
                                var charSlot = slot.Substring(i, 1);

                                writer.WriteLine($"{code}\t{charDay}{charSlot}");
                            }
                        }
                    }
                }
            }

            Console.WriteLine("Done");
        }
    }
}
