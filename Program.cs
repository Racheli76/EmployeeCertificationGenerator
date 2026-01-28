using EmployeeCertificationGenerator.Services;

// Initialize services
var loader = new CsvLoader();
var calculator = new ScoreCalculator();
var resolver = new CertificationResolver();

// Load employees from CSV
var csvPath = Path.Combine("Data", "Data.csv");
var employees = loader.Load(csvPath);

Console.WriteLine($"Loaded {employees.Count} employees from CSV\n");

// Process each employee
foreach (var employee in employees)
{
    // Calculate final score
    calculator.Calculate(employee);

    // Determine certification result
    var result = resolver.Resolve(employee.FinalScore);

    // Display result
    Console.WriteLine($"{employee.FullName} - Score: {employee.FinalScore:F2} - Status: {result}");
}

Console.WriteLine("\n✓ Processing complete");
