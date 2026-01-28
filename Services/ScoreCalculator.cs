using EmployeeCertificationGenerator.Models;

namespace EmployeeCertificationGenerator.Services
{
    /// <summary>
    /// Calculates the final score for an employee based on practical and theoretical scores.
    /// Formula: FinalScore = (PracticalScore * 0.6) + (TheoreticalScore * 0.4)
    /// </summary>
    public class ScoreCalculator
    {
        private const double PRACTICAL_WEIGHT = 0.6;
        private const double THEORETICAL_WEIGHT = 0.4;
        private const int SCORE_DECIMAL_PLACES = 2;

        /// <summary>
        /// Calculates and sets the final score for an employee.
        /// Result is rounded to 2 decimal places for clean document output.
        /// </summary>
        /// <param name="employee">The employee to calculate score for</param>
        /// <exception cref="ArgumentNullException">Thrown when employee is null</exception>
        public void Calculate(Employee employee)
        {
            if (employee == null)
                throw new ArgumentNullException(nameof(employee), "Employee cannot be null");

            var rawScore = (employee.PracticalScore * PRACTICAL_WEIGHT) +
                          (employee.TheoreticalScore * THEORETICAL_WEIGHT);

            employee.FinalScore = Math.Round(rawScore, SCORE_DECIMAL_PLACES);
        }
    }
}
