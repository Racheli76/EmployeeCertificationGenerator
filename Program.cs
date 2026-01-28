using EmployeeCertificationGenerator.Services;
using EmployeeCertificationGenerator.Models;

class Program
{
    static void Main(string[] args)
    {
        try
        {
            Console.WriteLine("Starting Employee Certification Generator...");
            Console.WriteLine();

            // Step 1: Load employees from CSV
            Console.WriteLine("Loading data from CSV...");
            var loader = new CsvLoader();
            var csvPath = Path.Combine("Data", "Data.csv");
            var employees = loader.Load(csvPath);
            Console.WriteLine($"Loaded {employees.Count} employees from CSV.");
            Console.WriteLine();

            // Step 2: Clean and normalize data
            Console.WriteLine("Cleaning and normalizing data...");
            var cleaner = new DataCleaner();
            employees = cleaner.Clean(employees);
            Console.WriteLine($"After cleaning: {employees.Count} employees.");
            Console.WriteLine();

            // Step 3: Calculate final scores
            Console.WriteLine("Calculating final scores...");
            var calculator = new ScoreCalculator();
            foreach (var employee in employees)
            {
                calculator.Calculate(employee);
            }
            Console.WriteLine("Final scores calculated for all employees.");
            Console.WriteLine();

            // Step 4: Generate PDF letters
            Console.WriteLine("Generating PDF letters...");
            var generator = new DocumentGenerator();
            generator.Generate(employees);
            Console.WriteLine("PDF generation completed.");
            Console.WriteLine();

            Console.WriteLine("✓ All processes completed successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }

        Console.WriteLine();
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
}
