# Employee Certification Generator

[![C#](https://img.shields.io/badge/Language-C%23-239120?logo=csharp&logoColor=white)](https://docs.microsoft.com/en-us/dotnet/csharp/)
[![.NET](https://img.shields.io/badge/Platform-.NET%206.0%2B-512BD4?logo=dotnet&logoColor=white)](https://dotnet.microsoft.com/)
[![Visual Studio](https://img.shields.io/badge/IDE-Visual%20Studio-5C2D91?logo=visualstudio&logoColor=white)](https://visualstudio.microsoft.com/)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

---

## Overview

**Employee Certification Generator** is a production-grade .NET 6+ console application that automates the complete certification workflow:

```
CSV Data  →  Clean  →  Calculate Scores  →  Generate PDFs  →  Archive Results
```

The system processes employee performance data, applies weighted scoring formulas, and generates personalized PDF certification letters with automated categorization (pass/fail/leadership track).

---

## Core Logic & Business Rules

### 🎯 Weighted Score Calculation

The final certification score applies a **60/40 weighted formula**:

```
Final Score = (Practical Score × 0.60) + (Theoretical Score × 0.40)
```

**Example:** Employee with Practical: 95, Theoretical: 80
- Calculation: (95 × 0.6) + (80 × 0.4) = 57 + 32 = **89.00**

### 📊 Certification Thresholds

| Score Range    | Status            | Output Action               |
|----------------|-------------------|-----------------------------|
| **< 70**       | **Failed**        | No document generated       |
| **70 – 89**    | **Passed**        | PDF with "לא עברת" notice   |
| **≥ 90**       | **Passed + Lead** | PDF with leadership track   |

#### The 90-Point Leadership Threshold

**Score ≥ 90 (Leadership Track Eligible)**
- Example: Rachel Green (95.20) receives:
  - ✅ Certification letter confirming pass
  - ✅ **Special designation:** "נמצאת מתאימ/ה לתפקיד מוביל/ה טכנולוגי מחלקתית"
  - ✅ Status: Leadership/Technological role eligibility

**Score 70-89 (Passed, Not Eligible)**
- Example: Yossi Israeli (89.00) receives:
  - ✅ Confirmation letter
  - ⚠️ Notice: "לא עברת את ההכשרה אך לצערנו לא נמצא תפקיד מתאים עבורך"
  - Status: Passed but ineligible for leadership roles

### 🔤 Missing Data Handling

When employee CSV data lacks Phone or Email fields, the system gracefully substitutes:

```
Missing Value  →  "לא הוזן" (Not Provided)
```

This ensures PDF documents remain complete while acknowledging data gaps.

---

## 🛠 Technical Stack

| Component                    | Purpose                          |
|------------------------------|----------------------------------|
| **🛠 C# / .NET 6.0+**        | Core application framework       |
| **📄 Xceed DocX**            | Word template manipulation       |
| **📑 Spire.Doc**             | DOCX → PDF conversion            |
| **🌐 CultureInfo.Invariant** | Culture-safe decimal parsing     |

### Technology Decisions

- **Line-by-Line CSV Reading:** Memory-efficient for large datasets
- **CultureInfo.InvariantCulture:** Robust decimal parsing across regional locales (prevents . vs , conflicts)
- **Mail Merge Notation (`« »`):** Uses guillemets for field replacement (Xceed DocX limitation workaround)
- **Explicit Resource Disposal:** Critical for Spire.Doc (must call `Close()` + `Dispose()`)

---

## Project Structure

```
EmployeeCertificationGenerator/
│
├── Data/
│   └── Data.csv                         # Employee source data
│
├── Models/
│   └── Employee.cs                      # Domain entity
│
├── Services/
│   ├── CsvLoader.cs                     # CSV parsing & mapping
│   ├── DataCleaner.cs                   # Normalization & deduplication
│   ├── ScoreCalculator.cs               # Weighted score computation
│   ├── CertificationResolver.cs         # Status determination
│   └── DocumentGenerator.cs             # PDF generation (Mail Merge)
│
├── Templates/
│   └── LetterTemplate.docx              # Word template with « » fields
│
├── Output/
│   ├── Passed/                          # PDF output (scores ≥ 70)
│   └── Failed/                          # Archive location
│
├── Program.cs                           # Pipeline orchestration
├── EmployeeCertificationGenerator.csproj
└── EmployeeCertificationGenerator.sln
```

---

## Pipeline Flow

### 1️⃣ **Load** (CsvLoader)
- Reads `Data/Data.csv` line-by-line
- Parses 5 required fields: `FirstName`, `LastName`, `Department`, `TheoreticalScore`, `PracticalScore`
- Uses `CultureInfo.InvariantCulture` for robust decimal parsing
- Gracefully skips malformed rows

**Input CSV Format:**
```csv
FirstName,LastName,Department,TheoreticalScore,PracticalScore
Dana,Cohen,Sales,80,95
Yossi,Israeli,IT,70,98.5
```

### 2️⃣ **Clean** (DataCleaner)
- **Normalizes names:** Converts to proper case (John, Doe, Sales)
- **Removes duplicates:** Case-insensitive detection using HashSet
- **Single-pass optimization:** Both operations in one loop

### 3️⃣ **Calculate** (ScoreCalculator)
- Applies weighted formula: `(Practical × 0.6) + (Theoretical × 0.4)`
- Rounds to 2 decimal places
- Sets `FinalScore` on each Employee object

### 4️⃣ **Resolve** (CertificationResolver)
- Categorizes into: Failed, Passed, or PassedExcellent
- Uses constants: `PASSING_SCORE = 70.0`, `EXCELLENCE_SCORE = 90.0`

### 5️⃣ **Generate** (DocumentGenerator)
- Loads `Templates/LetterTemplate.docx`
- Replaces Mail Merge fields:
  - `«FullName»` → "Dana Cohen"
  - `«Department»` → "Sales"
  - `«FinalScore»` → "89.00"
  - `«Phone»` → Value or "לא הוזן"
  - `«Email»` → Value or "לא הוזן"
  - `«BodyText»` → Score-dependent message
- Converts to PDF via Spire.Doc
- Saves to `Output/Passed/` or `Output/Failed/`

---

## Robustness & File Handling

### ✅ PDF File Locking Prevention

The DocumentGenerator implements defensive programming for file I/O:

```csharp
// Pre-delete existing PDFs to avoid "file in use" conflicts
if (File.Exists(pdfPath))
{
    File.Delete(pdfPath);
}
```

### ✅ Resource Disposal Strategy

**DocX (Xceed):**
```csharp
using (DocX document = DocX.Load(TemplatePath))
{
    // Automatic disposal via using block
}
```

**Spire.Doc:**
```csharp
finally
{
    spireDoc.Close();      // Explicitly close
    spireDoc.Dispose();    // Release file handle
}
```

### ✅ Error Handling

- Try-catch per document generation (isolation)
- Finally blocks ensure cleanup of temporary files
- User-friendly console output with exception messages

---

## Usage

### Prerequisites
- .NET 6.0+ runtime
- Visual Studio 2022+ (or any compatible IDE)

### Running the Application

```powershell
dotnet run
```

**Console Output:**
```
Starting Employee Certification Generator...

Loading data from CSV...
Loaded 9 employees from CSV.

Cleaning and normalizing data...
After cleaning: 9 employees.

Calculating final scores...
Final scores calculated for all employees.

Generating PDF letters...
Generated PDF: Output\Passed\Dana_Cohen_Certification.pdf
Generated PDF: Output\Passed\Yossi_Israeli_Certification.pdf
Generated PDF: Output\Passed\Rachel_Green_Certification.pdf

✓ All processes completed successfully!
```

---

## CSV Input Format

**File:** `Data/Data.csv`

| Column              | Type   | Example       | Required |
|---------------------|--------|---------------|----------|
| FirstName           | string | Dana          | Yes      |
| LastName            | string | Cohen         | Yes      |
| Department          | string | Sales         | Yes      |
| TheoreticalScore    | double | 80.0          | Yes      |
| PracticalScore      | double | 95.0          | Yes      |

**Notes:**
- Header row: Automatically skipped
- Phone/Email: Not in CSV; system uses "לא הוזן" fallback

---

## Business Rules Summary

| Rule                            | Value  | Source                   |
|---------------------------------|--------|--------------------------|
| Practical Weight                | 60%    | Business requirement     |
| Theoretical Weight              | 40%    | Business requirement     |
| Pass Threshold                  | 70     | CertificationResolver    |
| Leadership Track Threshold      | 90     | CertificationResolver    |
| Final Score Decimal Places      | 2      | ScoreCalculator          |
| Duplicate Detection             | Case-insensitive | DataCleaner |
| Missing Field Fallback          | "לא הוזן" | DocumentGenerator |

---

## Architecture Principles

✅ **Separation of Concerns:** Each service has a single responsibility
✅ **Dependency Injection Ready:** Services are stateless and testable
✅ **Error Resilience:** Graceful handling of malformed data
✅ **Resource Management:** Explicit disposal patterns for file I/O
✅ **Maintainability:** Constants centralize business rules (PASSING_SCORE, EXCELLENCE_SCORE, weights)

---

## Output Examples

### ✅ Leadership Track (Score ≥ 90)
```
הרינו להודיעך כי עברת בהצלחה את ההכשרה. 
הציון הסופי שלך הינו 95.20. 
נמצאת מתאימ/ה לתפקיד מוביל/ה טכנולוגי מחלקתית.
```

### ⚠️ Passed, Not Leadership (70-89)
```
הרינו להודיעך כי לא עברת את ההכשרה אך לצערנו לא נמצא תפקיד מתאים עבורך.
```

---

## License

MIT License – See [LICENSE](LICENSE) for details

---

## Author

Created as a professional homework assignment demonstrating enterprise-grade C# practices, data processing pipelines, and document generation workflows.

**Last Updated:** January 28, 2026