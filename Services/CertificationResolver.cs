namespace EmployeeCertificationGenerator.Services
{
    /// <summary>
    /// Certification outcome: Failed (less than 70), Passed (70-89), or PassedExcellent (90+, leadership track eligible).
    /// </summary>
    public enum CertificationResult
    {
        Failed,
        Passed,
        PassedExcellent
    }

    /// <summary>
    /// Resolves certification status based on final score using business-defined thresholds.
    /// </summary>
    public class CertificationResolver
    {
        /// <summary>Business requirement: Minimum score to pass certification (70 points).</summary>
        private const double PASSING_SCORE = 70.0;

        /// <summary>Business requirement: Minimum score for leadership track eligibility (90 points).</summary>
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
