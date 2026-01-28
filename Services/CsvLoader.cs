using EmployeeCertificationGenerator.Models;
using System.Globalization;

namespace EmployeeCertificationGenerator.Services
{
    /// <summary>
    /// Loads employee data from a CSV file and maps it to Employee objects.
    /// CSV format: FirstName,LastName,Department,TheoreticalScore,PracticalScore
    /// </summary>
    public class CsvLoader
    {
        /// <summary>
        /// Loads all employees from the CSV file.
        /// Skips the header row and parses each line into an Employee object.
        /// Uses line-by-line reading for memory efficiency and culture-invariant parsing for robustness.
        /// </summary>
        /// <param name="filePath">Path to the CSV file</param>
        /// <returns>List of Employee objects</returns>
        public List<Employee> Load(string filePath)
        {
            var employees = new List<Employee>();

            try
            {
                using (var reader = new StreamReader(filePath))
                {
                    string? line;
                    bool isHeader = true;

                    // Read file line-by-line for memory efficiency
                    while ((line = reader.ReadLine()) != null)
                    {
                        // Skip the header row
                        if (isHeader)
                        {
                            isHeader = false;
                            continue;
                        }

                        // Skip empty lines
                        if (string.IsNullOrWhiteSpace(line))
                            continue;

                        var parts = line.Split(',');

                        // Ensure we have exactly 5 fields
                        if (parts.Length != 5)
                            continue;

                        // Validate and parse scores with culture invariance and robust error handling
                        if (!double.TryParse(parts[3].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var theoreticalScore))
                            continue;

                        if (!double.TryParse(parts[4].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var practicalScore))
                            continue;

                        // Map CSV columns to Employee model
                        var employee = new Employee
                        {
                            FirstName = parts[0].Trim(),
                            LastName = parts[1].Trim(),
                            Department = parts[2].Trim(),
                            TheoreticalScore = theoreticalScore,
                            PracticalScore = practicalScore
                            // Email and Phone remain null - not provided in current data source
                        };

                        employees.Add(employee);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading CSV file: {ex.Message}");
            }

            return employees;
        }
    }
}
