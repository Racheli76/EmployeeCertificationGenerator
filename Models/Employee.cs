namespace EmployeeCertificationGenerator.Models
{

    public class Employee
    {
        // Data fields from CSV 
        public string FirstName { get; set; } = string.Empty;
        public string LastName { get; set; } = string.Empty;
        public string Department { get; set; } = string.Empty;
        public double TheoreticalScore { get; set; }
        public double PracticalScore { get; set; }

        // Business requirements - may be missing from initial data source
        public string? Email { get; set; }
        public string? Phone { get; set; }

        // Calculated final score - set once after calculation
        public double FinalScore { get; set; }

        // Convenience property for full name display
        public string FullName => $"{FirstName} {LastName}".Trim();
    }
}
