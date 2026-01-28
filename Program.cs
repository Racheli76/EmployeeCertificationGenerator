using EmployeeCertificationGenerator.Services;

// Initialize services
var loader = new CsvLoader();
var cleaner = new DataCleaner();
var calculator = new ScoreCalculator();
var resolver = new CertificationResolver();

// Load employees from CSV
var csvPath = Path.Combine("Data", "Data.csv");
var employees = loader.Load(csvPath);

Console.WriteLine($"Loaded {employees.Count} employees from CSV");

// Clean and normalize data
employees = cleaner.Clean(employees);

Console.WriteLine($"After cleaning: {employees.Count} employees\n");

// Process each employee
foreach (var employee in employees)
{
    // Calculate final score
    calculator.Calculate(employee);

    // Determine certification result
    var result = resolver.Resolve(employee.FinalScore);

    // Display result
    Console.WriteLine($"{employee.FullName} ({employee.Department}) - Score: {employee.FinalScore:F2} - Status: {result}");
}

Console.WriteLine("\n✓ Processing complete");
