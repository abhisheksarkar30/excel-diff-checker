import subprocess
import os
import pytest

# Define base directory and paths to your test files
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_DIFF_JAR_PATH = os.path.join(BASE_DIR, "excel-diff-checker-1/apps/excel-diff-checker.jar")
BASE_FILE_PATH = os.path.join(BASE_DIR, "Input/file.xlsx")
DIFF_FILE_PATH = os.path.join(BASE_DIR, "Input/file_diff.xlsx")
EMPTY_FILE_PATH = os.path.join(BASE_DIR, "Input/file_empty.xlsx")
IDENTICAL_FILE_PATH = os.path.join(BASE_DIR, "Input/file_identical.xlsx")
LARGE_FILE_PATH = os.path.join(BASE_DIR, "Input/large_file.xlsx")

# Helper function to construct the output file path
def get_output_file_path(base_file, target_file):
    base_name = os.path.splitext(os.path.basename(base_file))[0]
    target_name = os.path.splitext(os.path.basename(target_file))[0]
    output_file_name = f"{base_name} vs {target_name}.xlsx"
    return os.path.join(BASE_DIR, output_file_name)

# Helper function to run the excel-diff-checker command
def run_diff_checker(base_file, target_file, options=None):
    command = f"java -jar {EXCEL_DIFF_JAR_PATH} -b {base_file} -t {target_file}"
    if options:
        command += " " + options
    result = subprocess.run(command, shell=True, capture_output=True, text=True)
    return result.stdout, result.stderr

@pytest.fixture # Reference https://docs.pytest.org/en/stable/explanation/fixtures.html
def clean_up():
    yield
    # Clean up generated output files after tests
    for base_file in [BASE_FILE_PATH, DIFF_FILE_PATH, EMPTY_FILE_PATH, IDENTICAL_FILE_PATH, LARGE_FILE_PATH]:
        for target_file in [DIFF_FILE_PATH, EMPTY_FILE_PATH, IDENTICAL_FILE_PATH, LARGE_FILE_PATH]:
            output_file_path = get_output_file_path(base_file, target_file)
            if os.path.exists(output_file_path):
                os.remove(output_file_path)

# Test case F-1: Basic Functionality Test
def test_basic_functionality(clean_up):
    stdout, stderr = run_diff_checker(BASE_FILE_PATH, DIFF_FILE_PATH)
    OUTPUT_FILE_PATH = get_output_file_path(BASE_FILE_PATH, DIFF_FILE_PATH)
    assert os.path.exists(OUTPUT_FILE_PATH), "Output file was not generated."
    assert "Diff" in stdout, "No differences found in the output."
    assert "Diff at Row[3]" in stdout, "Expected specific difference in age not found."

# Test case F-2: Identical File Diff Check
def test_identical_files(clean_up):
    stdout, stderr = run_diff_checker(BASE_FILE_PATH, IDENTICAL_FILE_PATH)
    OUTPUT_FILE_PATH = get_output_file_path(BASE_FILE_PATH, IDENTICAL_FILE_PATH)
    assert not os.path.exists(OUTPUT_FILE_PATH), "No output file should be generated for identical files."
    assert "No diff found!" in stdout, "Unexpected output for identical files."

# Test case F-3: Diffs in Comments/Notes
def test_diffs_in_comments(clean_up):
    stdout, stderr = run_diff_checker(BASE_FILE_PATH, DIFF_FILE_PATH)
    assert "Diff at Row[3] of Sheet[Sheet1]" in stdout, "Differences not correctly reported in comments."

# Test case F-4: Cell Difference Detection
def test_cell_difference_detection(clean_up):
    stdout, stderr = run_diff_checker(BASE_FILE_PATH, DIFF_FILE_PATH)
    assert "Diff at Row[3]" in stdout, "Cell differences were not detected properly."

# Test case F-5: Empty File Comparison
def test_empty_file_comparison(clean_up):
    stdout, stderr = run_diff_checker(EMPTY_FILE_PATH, EMPTY_FILE_PATH)
    OUTPUT_FILE_PATH = get_output_file_path(EMPTY_FILE_PATH, EMPTY_FILE_PATH)
    assert not os.path.exists(OUTPUT_FILE_PATH), "No output file should be generated for empty files."
    assert "No diff found!" in stdout, "Unexpected output for empty file comparison."

# Test case F-6: Specific Sheet Comparison (Option -s)
def test_specific_sheet_comparison(clean_up):
    options = "-s Sheet1"
    stdout, stderr = run_diff_checker(BASE_FILE_PATH, DIFF_FILE_PATH, options)
    OUTPUT_FILE_PATH = get_output_file_path(BASE_FILE_PATH, DIFF_FILE_PATH)
    assert os.path.exists(OUTPUT_FILE_PATH), "Output file was not generated for specific sheet comparison."
    assert "Diff" in stdout, "No differences found in the output for specific sheet comparison."
    assert "Diff at Row[3]" in stdout, "Expected specific difference in age not found."

# Test case F-7: Mixed Data Types
def test_mixed_data_types(clean_up):
    stdout, stderr = run_diff_checker(BASE_FILE_PATH, DIFF_FILE_PATH)
    OUTPUT_FILE_PATH = get_output_file_path(BASE_FILE_PATH, DIFF_FILE_PATH)
    assert os.path.exists(OUTPUT_FILE_PATH), "Output file was not generated for mixed data type comparison."
    assert "Diff" in stdout, "No differences found in the output for mixed data types."
    assert "Diff at Row[3]" in stdout, "Expected name not found in the differences."

# Test case F-8: Redundant Row Removal
def test_redundant_row_removal(clean_up):
    stdout, stderr = run_diff_checker(BASE_FILE_PATH, DIFF_FILE_PATH)
    assert "Diff at Row" in stdout, "Redundant row differences not handled properly."

# Test case F-9: Large File Comparison
def test_large_file_comparison(clean_up):
    stdout, stderr = run_diff_checker(BASE_FILE_PATH, LARGE_FILE_PATH)
    OUTPUT_FILE_PATH = get_output_file_path(BASE_FILE_PATH, LARGE_FILE_PATH)
    assert os.path.exists(OUTPUT_FILE_PATH), "Output file was not generated for large file comparison."
    assert "Diff" in stdout, "No differences found in the output for large file comparison."

# Test case F-10: Use of 'r' Option
def test_use_of_r_option(clean_up):
    options = "-r"
    stdout, stderr = run_diff_checker(BASE_FILE_PATH, DIFF_FILE_PATH, options)
    assert "Diff at Cell" in stdout, "Differences not reported correctly when using 'r' option."
    assert not os.path.exists(get_output_file_path(BASE_FILE_PATH, DIFF_FILE_PATH)), "No output file should be generated with 'r' option."

# Test case F-11: Non-Existent Sheet Name
def test_non_existent_sheet_name(clean_up):
    options = "-s NonExistentSheet"
    stdout, stderr = run_diff_checker(BASE_FILE_PATH, DIFF_FILE_PATH, options)
    assert "doesn't exist in both workbooks" in stdout, "Non-existent sheet name was not handled properly."

# Test case F-12: Invalid File Path
def test_file_not_found(clean_up):
    stdout, stderr = run_diff_checker("invalid_path.xlsx", DIFF_FILE_PATH)
    assert "The system cannot find the file specified" in stdout or "The system cannot find the file specified" in stderr, "No error message for missing file."