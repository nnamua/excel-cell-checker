from openpyxl.styles import PatternFill
from openpyxl.cell.cell import Cell
from openpyxl.comments import Comment
import violations

"""
Utility functions and variables.
"""

# Checks the JSON structure file
REQUIRED_KEYS = ( "name", "type" )
def check_struct(struct):
    if "cols" not in struct:
        return "'cols' key is missing."
    else:
        cols = struct["cols"]
        for index, col in enumerate(cols):
            for required_key in REQUIRED_KEYS:
                if col.get("skip", True):
                    continue
                elif required_key not in REQUIRED_KEYS:
                    return f"Key '{required_key}' is required for all unskipped columns, but was not found for column #{index}."
    return None

# The following functions add ANSI-escape sequences for text highlighting

# Green
def g(string):
    return f"\033[32m{string}\033[0m"

# Red
def r(string):
    return f"\033[31m{string}\033[0m"

# Yellow
def y(string):
    return f"\033[33m{string}\033[0m"

# Underline
def ul(string):
    return f"\033[4m{string}\033[0m"

# Bold
def b(string):
    return f"\033[1m{string}\033[0m"

# Prints with indentation
def print_indent(string, indent=0):
    string = string.replace("\n", "\n"+" "*indent)
    print(f"{' ' * indent}{string}")

# Summary colored text
OK    = f"[{g('OK')}]"
ERROR = f"[{r('ERROR')}]"
SKIPPED = f"[{y('SKIPPED')}]"

# Red background fill
red_fill = PatternFill(
    start_color="ffff9797",
    end_color="ffff9797",
    fill_type="solid"
)

# Marks the given cell
def mark(cell, violation):
    if not isinstance(cell, Cell):
        return
    
    cell.fill = red_fill

    if isinstance(violation, violations.NonEmptyViolation):
        text = "Cell must not be empty."
    elif isinstance(violation, violations.TypeViolation):
        text = f"Cell is of type '{violation.actual.__name__}', but should have been '{violation.expected.__name__}'."
    elif isinstance(violation, violations.RegexViolation):
        text = f"Cell does not match the specified regular expression '{violation.pattern}'."
    else:
        # unknown violation
        return

    comment = Comment(text, "excel-cell-checker")
    cell.comment = comment