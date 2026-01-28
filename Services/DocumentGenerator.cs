using EmployeeCertificationGenerator.Models;
using Xceed.Words.NET;
using Spire.Doc;
using System.Globalization;

namespace EmployeeCertificationGenerator.Services
{
    /// <summary>
    /// Generates PDF certification letters from Word templates.
    /// Uses Xceed.Words.NET for Word manipulation and Spire.Doc for PDF conversion.
    /// </summary>
    public class DocumentGenerator
    {
        private const string TemplatePath = "Templates/LetterTemplate.docx";
        private const string OutputFolder = "Output";

        /// <summary>
        /// Generates PDF certification letters for all eligible employees (FinalScore >= 70).
        /// </summary>
        /// <param name="employees">List of employees to process</param>
        /// <exception cref="ArgumentNullException">Thrown when employees list is null</exception>
        public void Generate(List<Employee> employees)
        {
            if (employees == null)
                throw new ArgumentNullException(nameof(employees), "Employee list cannot be null");

            // Ensure output folder exists
            if (!Directory.Exists(OutputFolder))
            {
                Directory.CreateDirectory(OutputFolder);
            }

            foreach (var employee in employees)
            {
                try
                {
                    // Only generate for employees who passed (FinalScore >= 70)
                    if (employee.FinalScore < 70)
                        continue;

                    GenerateSingleDocument(employee);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error generating document for {employee.FullName}: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Generates a single PDF document for an employee.
        /// </summary>
        private void GenerateSingleDocument(Employee employee)
        {
            // Create temporary Word file path
            string tempDocxPath = Path.Combine(OutputFolder, $"{employee.FullName}_temp.docx");
            string pdfPath = Path.Combine(OutputFolder, $"{employee.FullName}_Certification.pdf");

            try
            {
                // Load template and replace placeholders
                using (DocX document = DocX.Load(TemplatePath))
                {
                    // Replace placeholders
                    document.ReplaceText("{{FullName}}", employee.FirstName + " " + employee.LastName);
                    document.ReplaceText("{{Department}}", employee.Department);
                    document.ReplaceText("{{Phone}}", string.IsNullOrWhiteSpace(employee.Phone) ? "לא הוזן" : employee.Phone);
                    document.ReplaceText("{{Email}}", string.IsNullOrWhiteSpace(employee.Email) ? "לא הוזן" : employee.Email);
                    document.ReplaceText("{{FinalScore}}", employee.FinalScore.ToString(CultureInfo.InvariantCulture));

                    // Determine body text based on score
                    string bodyText = GetBodyText(employee);
                    document.ReplaceText("{{BodyText}}", bodyText);

                    // Save temporary Word document
                    document.SaveAs(tempDocxPath);
                }

                // Convert to PDF using Spire.Doc
                Document spireDoc = new Document();
                spireDoc.LoadFromFile(tempDocxPath);
                spireDoc.SaveToFile(pdfPath, Spire.Doc.FileFormat.PDF);

                Console.WriteLine($"Generated PDF: {pdfPath}");
            }
            finally
            {
                // Clean up temporary file
                if (File.Exists(tempDocxPath))
                {
                    File.Delete(tempDocxPath);
                }
            }
        }

        /// <summary>
        /// Gets the appropriate body text based on the employee's final score.
        /// </summary>
        private string GetBodyText(Employee employee)
        {
            if (employee.FinalScore >= 90)
            {
                return $"הרינו להודיעך כי עברת בהצלחה את ההכשרה. הציון הסופי שלך הינו {employee.FinalScore}. נמצאת מתאימ/ה לתפקיד מוביל/ה טכנולוגי מחלקתית.";
            }
            else if (employee.FinalScore >= 70)
            {
                return $"הרינו להודיעך כי עברת בהצלחה את ההכשרה. הציון הסופי שלך הינו {employee.FinalScore}.";
            }
            else
            {
                // This should not be reached since we filter >=70, but for completeness
                return "הרינו להודיעך כי לא עברת את ההכשרה אך לצערנו לא נמצא תפקיד מתאים עבורך.";
            }
        }
    }
}