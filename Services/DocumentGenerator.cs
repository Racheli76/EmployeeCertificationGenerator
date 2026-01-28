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
        /// Generates a single PDF document for an employee using Mail Merge field placeholders.
        /// Replaces placeholders in the format «FieldName» (guillemets notation).
        /// Includes robust resource management and file locking prevention.
        /// </summary>
        private void GenerateSingleDocument(Employee employee)
        {
            // Create temporary Word file path
            string tempDocxPath = Path.Combine(OutputFolder, $"{employee.FullName}_temp.docx");
            string pdfPath = Path.Combine(OutputFolder, $"{employee.FullName}_Certification.pdf");

            try
            {
                // Step 1: Delete existing PDF if it exists (to avoid "file in use" errors)
                if (File.Exists(pdfPath))
                {
                    try
                    {
                        File.Delete(pdfPath);
                    }
                    catch (IOException ex)
                    {
                        Console.WriteLine($"Warning: Could not delete existing PDF '{pdfPath}'. File may be open in another program: {ex.Message}");
                        return;
                    }
                }

                // Step 2: Load template and replace placeholders
                using (DocX document = DocX.Load(TemplatePath))
                {
                    // Replace Mail Merge field placeholders (using « » guillemets notation)
                    document.ReplaceText("«FullName»", employee.FirstName + " " + employee.LastName);
                    document.ReplaceText("«Department»", employee.Department);
                    document.ReplaceText("«Phone»", string.IsNullOrWhiteSpace(employee.Phone) ? "לא הוזן" : employee.Phone);
                    document.ReplaceText("«Email»", string.IsNullOrWhiteSpace(employee.Email) ? "לא הוזן" : employee.Email);
                    document.ReplaceText("«FinalScore»", employee.FinalScore.ToString("F1", CultureInfo.InvariantCulture));
                    document.ReplaceText("«BodyText»", GetBodyText(employee));

                    // Save temporary Word document
                    document.SaveAs(tempDocxPath);
                } // DocX document is disposed here, releasing file handle

                // Step 3: Convert temporary DOCX to PDF using Spire.Doc
                // Ensure the temporary file is fully closed before Spire attempts to load it
                Document spireDoc = new Document();
                try
                {
                    spireDoc.LoadFromFile(tempDocxPath);
                    spireDoc.SaveToFile(pdfPath, Spire.Doc.FileFormat.PDF);
                    Console.WriteLine($"Generated PDF: {pdfPath}");
                }
                finally
                {
                    // Explicitly close and dispose Spire.Doc to release file handle
                    spireDoc.Close();
                    spireDoc.Dispose();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating document for {employee.FullName}: {ex.Message}");
            }
            finally
            {
                // Step 4: Clean up temporary file
                if (File.Exists(tempDocxPath))
                {
                    try
                    {
                        File.Delete(tempDocxPath);
                    }
                    catch (IOException ex)
                    {
                        Console.WriteLine($"Warning: Could not delete temporary file '{tempDocxPath}': {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// Gets the appropriate body text based on the employee's final score.
        /// Strictly follows assignment requirements:
        /// - FinalScore >= 90: Passed with leadership role text
        /// - 70 <= FinalScore < 90: Standard rejection text
        /// - FinalScore < 70: No document generated
        /// </summary>
        private string GetBodyText(Employee employee)
        {
            if (employee.FinalScore >= 90)
            {
                return $"הרינו להודיעך כי עברת בהצלחה את ההכשרה. הציון הסופי שלך הינו {employee.FinalScore.ToString("F1", CultureInfo.InvariantCulture)}. נמצאת מתאימ/ה לתפקיד מוביל/ה טכנולוגי מחלקתית.";
            }
            else if (employee.FinalScore >= 70)
            {
                return "הרינו להודיעך כי לא עברת את ההכשרה אך לצערנו לא נמצא תפקיד מתאים עבורך.";
            }
            else
            {
                // This should not be reached since we filter <70, but for completeness
                return "הרינו להודיעך כי לא עברת את ההכשרה אך לצערנו לא נמצא תפקיד מתאים עבורך.";
            }
        }
    }
}