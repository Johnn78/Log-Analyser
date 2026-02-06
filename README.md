# Log Analysis Tool

Log analysis utility designed to process very large log files and extract
security, access and system-related events in a concise and readable format.

The tool was developed to support real production incident investigations,
where manual log inspection was not feasible due to file size and time constraints.

---

## Problem

In real production environments, log files may contain tens of thousands of lines.
Manually identifying meaningful events is time-consuming and error-prone.

Typical investigation targets include:
- Successful user logins
- Failed authentication attempts
- Suspicious access patterns
- User actions during disputed incidents

---

## Solution

This tool filters log files based on predefined keywords and produces a reduced
output containing only relevant log entries.

### Processing Flow
- **Input**
  - Log file (`.log`, `.txt`)
  - List of keywords (e.g. `logged in`, `login failed`)
- **Processing**
  - Line-by-line file scanning
  - Keyword matching
- **Output**
  - New text file with only matching entries

Large log files (e.g. 70,000+ lines) are reduced to a small, focused output suitable
for fast analysis.

---

## Usage

1. Open the application
2. Select the log file to analyze
3. Define the list of keywords
4. Run the analysis
5. Review the generated output file

---

## Real-World Usage

The tool has been used in real production investigations, including:
- Security access analysis
- Investigation of disputed financial transactions
- Verification of user-initiated actions versus software faults

In these cases, the tool significantly reduced investigation time and provided
clear technical evidence for incident resolution.

---

## Example

**Input**
- Log file: ~75,000 lines
- Keywords:
  - `logged in`
  - `logged successfully`
  - `login failed`

**Output**
- Filtered text file with ~50 relevant log entries

---

## Technologies

- Visual Basic 6  
- File I/O processing  
- Text parsing and keyword-based filtering  

> Visual Basic 6 was used due to production environment constraints at the time
> the tool was developed.

---

## Limitations

- Keyword-based filtering (no regex or semantic analysis)
- Designed for plain-text log files
- Windows environment

---

## License

This project is provided for demonstration and portfolio purposes.
