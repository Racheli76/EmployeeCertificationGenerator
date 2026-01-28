namespace EmployeeCertificationGenerator.Services
{
    /// <summary>
    /// Represents the outcome of an employee's certification.
    /// </summary>
    public enum CertificationResult
    {
        Failed,           // Score < 70
        Passed,           // Score 70-89
        PassedExcellent   // Score >= 90
    }

    /// <summary>
    /// Determines the certification result based on final score.
    /// Business rules:
    /// - Failed: < 70
    /// - Passed: 70-89
    /// - PassedExcellent: >= 90
    /// </summary>
    public class CertificationResolver
    {
        // Business rule thresholds - centralized for easy maintenance
        private const double PASSING_SCORE = 70.0;
        private const double EXCELLENCE_SCORE = 90.0;

        /// <summary>
        /// Resolves the certification result for a given score.
        /// Business rules:
        /// - Failed: score < 70
        /// - Passed: 70 <= score < 90
        /// - PassedExcellent: score >= 90
        /// </summary>
        /// <param name="finalScore">The final score to evaluate</param>
        /// <returns>The certification result</returns>
        public CertificationResult Resolve(double finalScore)
        {
            if (finalScore < PASSING_SCORE)
                return CertificationResult.Failed;

            if (finalScore >= EXCELLENCE_SCORE)
                return CertificationResult.PassedExcellent;

            return CertificationResult.Passed;
        }
    }
}
