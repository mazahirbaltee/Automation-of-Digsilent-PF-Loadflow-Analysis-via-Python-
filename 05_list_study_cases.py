"""
============================================================
05_list_study_cases.py

PowerFactory Python Toolkit

Purpose:
    List all Study Cases in the active project.

Author:
    Mazahir Hossein Learning Series

PowerFactory:
    2024

Python:
    3.11
============================================================
"""

# ----------------------------------------------------------
# Imports
# ----------------------------------------------------------

import sys

# ----------------------------------------------------------
# Configuration
# ----------------------------------------------------------

PF_PATH = r"C:\Program Files\DIgSILENT\PowerFactory 2024\Python\3.11"

PROJECT_NAME = "Python_Integration_Proj1"

if PF_PATH not in sys.path:
    sys.path.append(PF_PATH)

try:
    import powerfactory
except ImportError:
    print("ERROR : Unable to import PowerFactory API.")
    raise SystemExit()

# ----------------------------------------------------------
# Functions
# ----------------------------------------------------------

def connect_pf():

    app = powerfactory.GetApplicationExt()

    if app is None:
        raise RuntimeError("Unable to connect to PowerFactory.")

    return app


# ----------------------------------------------------------
# Main
# ----------------------------------------------------------

def main():

    print("=" * 60)
    print("POWERFACTORY STUDY CASE LIST")
    print("=" * 60)

    try:

        app = connect_pf()

        # Activate project first
        result = app.ActivateProject(PROJECT_NAME)

        if result != 0:
            print("Project activation failed.")
            return

        # Get Study Cases folder
        study_folder = app.GetProjectFolder("study")

        if study_folder is None:
            print("Study Cases folder not found.")
            return

        print(f"Project : {PROJECT_NAME}")
        print(f"Folder  : {study_folder.loc_name}")
        print()

        # Get all Study Cases
        study_cases = study_folder.GetContents("*.IntCase", 1)

        if not study_cases:
            print("No Study Cases found.")
            return

        print("Available Study Cases")
        print("-" * 60)

        for index, study_case in enumerate(study_cases, start=1):
            print(f"{index:2d}. {study_case.loc_name}")

        print("-" * 60)
        print(f"Total Study Cases : {len(study_cases)}")

    except Exception as err:

        print("ERROR")
        print(err)

    print("=" * 60)


# ----------------------------------------------------------
# Entry Point
# ----------------------------------------------------------

if __name__ == "__main__":
    main()