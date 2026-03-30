"""Script generator module: generates and executes Python scripts based on user instructions."""

import re
import logging
import tempfile
import traceback
from pathlib import Path

import pandas as pd

from backend.llm_client import call_llm

logger = logging.getLogger(__name__)

# Dangerous patterns that should not appear in generated scripts
DANGEROUS_PATTERNS = [
    r"\bos\b\s*\.\s*system",
    r"\bos\b\s*\.\s*popen",
    r"\bsubprocess\b",
    r"\b__import__\b",
    r"\beval\s*\(",
    r"\bcompile\s*\(",
    r"\bshutil\b\s*\.\s*rmtree",
    r"\bos\b\s*\.\s*remove",
    r"\bos\b\s*\.\s*unlink",
    r"\bopen\s*\([^)]*['\"]w['\"]",
]

UPLOADS_DIR = Path("uploads")
RESULTS_DIR = Path("results")


def get_file_info(filepath: Path) -> dict:
    """Collect metadata and preview data from an Excel/CSV file."""
    try:
        if filepath.suffix == ".csv":
            df = pd.read_csv(filepath, nrows=10)
            df_full = pd.read_csv(filepath)
        else:
            df = pd.read_excel(filepath, nrows=10)
            df_full = pd.read_excel(filepath)

        return {
            "filename": filepath.name,
            "rows": len(df_full),
            "columns": list(df_full.columns),
            "dtypes": {col: str(dtype) for col, dtype in df_full.dtypes.items()},
            "preview": df.head(5).to_dict(orient="records"),
        }
    except Exception as e:
        logger.error("Error reading file %s: %s", filepath, e)
        return {"filename": filepath.name, "error": str(e)}


def get_functions_info() -> str:
    """Return a description of available excel_functions for the LLM prompt."""
    return """
Available Excel processing functions (import from excel_functions):

### CRUD Operations (from excel_functions.crud):
- excel_add_row(filepath, sheet_name, row_data) - Add a new row
- excel_add_column(filepath, sheet_name, column_name, data=None, default_value=None) - Add a new column
- excel_delete_row(filepath, sheet_name, condition=None, row_index=None) - Delete rows
- excel_delete_column(filepath, sheet_name, column_name) - Delete columns
- excel_delete_empty_rows(filepath, sheet_name) - Remove empty rows
- excel_modify_cell(filepath, sheet_name, row_index, column_name, value) - Modify a cell
- excel_modify_column(filepath, sheet_name, column_name, condition, new_value) - Batch modify

### Query Operations (from excel_functions.query):
- excel_query_data(filepath, sheet_name, conditions) - Query by conditions
- excel_filter_by_value(filepath, sheet_name, column_name, values) - Filter by values
- excel_search_text(filepath, sheet_name, keyword) - Full-text search
- excel_get_unique_values(filepath, sheet_name, column_name) - Get unique values
- excel_filter_by_range(filepath, sheet_name, column_name, min_val, max_val) - Filter by range

### Statistics (from excel_functions.statistics):
- excel_sum_column(filepath, sheet_name, column_name) - Sum
- excel_average_column(filepath, sheet_name, column_name) - Average
- excel_count_values(filepath, sheet_name, column_name) - Count
- excel_max_value(filepath, sheet_name, column_name) - Max
- excel_min_value(filepath, sheet_name, column_name) - Min
- excel_deduplicate(filepath, sheet_name, columns) - Deduplicate
- excel_group_statistics(filepath, sheet_name, group_column, stat_column, stat_type) - Group stats
- excel_calculate_statistics(filepath, sheet_name, column_name) - Full statistics

### Merge Operations (from excel_functions.merge):
- excel_merge_files(file_list, output_file, merge_type) - Merge files
- excel_merge_sheets(filepath, output_file) - Merge sheets
- excel_join_files(file1, file2, output_file, on_column, join_type) - Join files
- excel_append_data(source_file, target_file, output_file) - Append data
"""


def validate_script(script: str) -> tuple[bool, str]:
    """
    Validate a generated script for security risks.

    Returns:
        Tuple of (is_safe, reason).
    """
    for pattern in DANGEROUS_PATTERNS:
        if re.search(pattern, script):
            return False, f"Dangerous pattern detected: {pattern}"

    return True, "Script is safe"


def build_prompt(file_info: list[dict], user_instruction: str) -> list[dict]:
    """Build the LLM prompt with file context and user instruction."""
    file_context = ""
    for info in file_info:
        if "error" in info:
            file_context += f"\nFile: {info['filename']} - Error: {info['error']}\n"
        else:
            file_context += f"\nFile: {info['filename']}\n"
            file_context += f"  Rows: {info['rows']}\n"
            file_context += f"  Columns: {info['columns']}\n"
            file_context += f"  Data types: {info['dtypes']}\n"
            preview_df = pd.DataFrame(info["preview"])
            file_context += f"  Preview:\n{preview_df.to_string(index=False)}\n"

    functions_info = get_functions_info()

    system_prompt = f"""You are an expert Python programmer specializing in Excel data processing.
You generate Python scripts to process Excel files based on user instructions.

IMPORTANT RULES:
1. The uploaded files are located in the "uploads/" directory.
2. Save all output files to the "results/" directory with a descriptive filename.
3. You can use pandas, openpyxl, and the excel_functions module.
4. If the task requires calling an LLM API (e.g., generating text, looking up information, etc.),
   use the helper function provided below. Do NOT use os, subprocess, or any system commands.
5. Return ONLY the Python code, wrapped in ```python ... ``` markers.
6. The script must be self-contained and executable.
7. Print a brief summary of what was done at the end.
8. For LLM calls, use this async pattern:

```python
import asyncio
from backend.llm_client import call_llm

async def ask_llm(question: str) -> str:
    messages = [{{"role": "user", "content": question}}]
    return await call_llm(messages)

# To call synchronously within the script:
# result = asyncio.run(ask_llm("your question"))
```

Available files:
{file_context}

{functions_info}
"""

    return [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_instruction},
    ]


def extract_code(response: str) -> str:
    """Extract Python code from LLM response."""
    # Try to extract from markdown code blocks
    pattern = r"```python\s*\n(.*?)```"
    matches = re.findall(pattern, response, re.DOTALL)
    if matches:
        return matches[0].strip()

    # Try generic code blocks
    pattern = r"```\s*\n(.*?)```"
    matches = re.findall(pattern, response, re.DOTALL)
    if matches:
        return matches[0].strip()

    # Return the full response as a fallback
    return response.strip()


async def generate_and_execute(
    uploaded_files: list[str], user_instruction: str
) -> dict:
    """
    Generate a Python script from user instruction, validate it, and execute it.

    Args:
        uploaded_files: List of uploaded filenames.
        user_instruction: Natural language instruction from the user.

    Returns:
        Dict with keys: success, message, script, output_files, preview
    """
    # Gather file info
    file_info = []
    for filename in uploaded_files:
        filepath = UPLOADS_DIR / filename
        if filepath.exists():
            file_info.append(get_file_info(filepath))

    if not file_info:
        return {
            "success": False,
            "message": "No valid uploaded files found.",
            "script": "",
            "output_files": [],
            "preview": None,
        }

    # Build prompt and call LLM
    messages = build_prompt(file_info, user_instruction)

    try:
        response = await call_llm(messages)
    except ValueError as e:
        return {
            "success": False,
            "message": str(e),
            "script": "",
            "output_files": [],
            "preview": None,
        }
    except Exception as e:
        logger.error("LLM API error: %s", e)
        return {
            "success": False,
            "message": f"LLM API call failed: {e}",
            "script": "",
            "output_files": [],
            "preview": None,
        }

    # Extract and validate script
    script = extract_code(response)
    is_safe, reason = validate_script(script)

    if not is_safe:
        return {
            "success": False,
            "message": f"Generated script failed security check: {reason}",
            "script": script,
            "output_files": [],
            "preview": None,
        }

    # Execute script
    RESULTS_DIR.mkdir(exist_ok=True)

    # Get list of result files before execution
    existing_results = set(RESULTS_DIR.iterdir()) if RESULTS_DIR.exists() else set()

    try:
        # Write script to a temp file and execute
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".py", dir="temp", delete=False
        ) as f:
            f.write(script)
            temp_script_path = f.name

        # Execute the script
        exec_globals = {"__builtins__": __builtins__}
        exec(compile(script, temp_script_path, "exec"), exec_globals)

        # Clean up temp file
        Path(temp_script_path).unlink(missing_ok=True)

    except Exception as e:
        logger.error("Script execution error: %s\n%s", e, traceback.format_exc())
        Path(temp_script_path).unlink(missing_ok=True)
        return {
            "success": False,
            "message": f"Script execution failed: {e}",
            "script": script,
            "output_files": [],
            "preview": None,
        }

    # Find new result files
    current_results = set(RESULTS_DIR.iterdir()) if RESULTS_DIR.exists() else set()
    new_files = current_results - existing_results
    output_filenames = [f.name for f in new_files]

    # Generate preview of the first result file
    preview = None
    if output_filenames:
        result_path = RESULTS_DIR / output_filenames[0]
        try:
            if result_path.suffix == ".csv":
                df = pd.read_csv(result_path)
            else:
                df = pd.read_excel(result_path)
            preview = {
                "filename": output_filenames[0],
                "columns": list(df.columns),
                "rows": len(df),
                "data": df.head(50).fillna("").to_dict(orient="records"),
            }
        except Exception as e:
            logger.warning("Could not preview result file: %s", e)

    return {
        "success": True,
        "message": "Processing completed successfully.",
        "script": script,
        "output_files": output_filenames,
        "preview": preview,
    }
