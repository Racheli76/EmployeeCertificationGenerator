using EmployeeCertificationGenerator.Models;

namespace EmployeeCertificationGenerator.Services
{
    /// <summary>
    /// Cleans and normalizes employee data.
    /// Handles name normalization, duplicate removal, and data consistency.
    /// </summary>
    public class DataCleaner
    {
        /// <summary>
        /// Cleans a list of employees by normalizing names and removing duplicates.
        /// Performs both operations in a single pass for memory efficiency.
        /// </summary>
        /// <param name="employees">List of employees to clean</param>
        /// <returns>Cleaned list with normalized names and no duplicates</returns>
        /// <exception cref="ArgumentNullException">Thrown when employee list is null</exception>
        public List<Employee> Clean(List<Employee> employees)
        {
            if (employees == null)
                throw new ArgumentNullException(nameof(employees), "Employee list cannot be null");

            var seen = new HashSet<string>();
            var cleaned = new List<Employee>();

            foreach (var employee in employees)
            {
                // Step 1: Normalize employee data first
                NormalizeEmployee(employee);

                // Step 2: Create unique key for duplicate detection
                // Using ToLower() for case-insensitive comparison ensures robustness
                var key = $"{employee.FirstName.ToLower()}_{employee.LastName.ToLower()}_{employee.Department.ToLower()}";

                // Step 3: Only add if not seen before
                if (!seen.Contains(key))
                {
                    seen.Add(key);
                    cleaned.Add(employee);
                }
            }

            return cleaned;
        }

        /// <summary>
        /// Normalizes all name fields for a single employee.
        /// </summary>
        private void NormalizeEmployee(Employee employee)
        {
            employee.FirstName = NormalizeName(employee.FirstName);
            employee.LastName = NormalizeName(employee.LastName);
            employee.Department = NormalizeDepartment(employee.Department);
        }

        /// <summary>
        /// Converts a name to proper case (first letter uppercase, rest lowercase).
        /// Example: "JOHN" → "John", "john" → "John", "A" → "A", "jOhN" → "John"
        /// </summary>
        private string NormalizeName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return string.Empty;

            // Trim whitespace
            name = name.Trim();

            // Handle empty result after trimming
            if (name.Length == 0)
                return string.Empty;

            // Handle single character - just uppercase it
            if (name.Length == 1)
                return name.ToUpper();

            // Safe to use Substring(1) since length is at least 2
            return char.ToUpper(name[0]) + name.Substring(1).ToLower();
        }

        /// <summary>
        /// Normalizes department name to proper case.
        /// </summary>
        private string NormalizeDepartment(string department)
        {
            if (string.IsNullOrWhiteSpace(department))
                return "Unknown";

            department = department.Trim();
            if (department.Length == 0)
                return "Unknown";

            return char.ToUpper(department[0]) + department.Substring(1).ToLower();
        }
    }
}