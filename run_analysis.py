#!/usr/bin/env python3
import sys
from analysis import analyze
from report_writer import write_weekly_report

if len(sys.argv) != 4:
    print("Usage: python run_analysis.py <raw.xlsx> <progress.xlsx> <gms.csv>")
    sys.exit(1)

raw_file, progress_file, gms_file = sys.argv[1:4]

# Run analysis
result = analyze(raw_file, progress_file, gms_file, progress_file)

# Generate report
output_path = f"Weekly_Top100_Report_Week_{result['week']}.docx"
write_weekly_report(result, output_path)

print(f"Report generated: {output_path}")