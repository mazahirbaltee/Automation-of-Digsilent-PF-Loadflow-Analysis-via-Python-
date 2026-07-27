"""
============================================================
06_activate_study_case.py

PowerFactory Python Toolkit

Purpose:
    Activate a Study Case and verify that
    the activation was successful.

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
STUDY_CASE_NAME = "Study Case"

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
    print("POWERFACTORY STUDY CASE ACTIVATION")
    print("=" * 60)

    try:

        app = connect_pf()

        # Activate Project
        result = app.ActivateProject(PROJECT_NAME)

        if result != 0:
            print("Project activation failed.")
            return

        # Get Study Case Folder
        study_folder = app.GetProjectFolder("study")

        if study_folder is None:
            print("Study Cases folder not found.")
            return

        # Find Study Case
        study_cases = study_folder.GetContents("*.IntCase", 1)

        selected_case = None

        for case in study_cases:
            if case.loc_name == STUDY_CASE_NAME:
                selected_case = case
                break

        if selected_case is None:
            print(f"Study Case '{STUDY_CASE_NAME}' not found.")
            return

        # Activate Study Case
        result = selected_case.Activate()

        if result != 1:
            print("Study Case activation failed.")
            print(f"Return Code : {result}")
            return

        # Verify
        active_case = app.GetActiveStudyCase()

        print(f"Project          : {PROJECT_NAME}")
        print(f"Requested Study  : {STUDY_CASE_NAME}")
        print()
        print("Status           : SUCCESS")
        print(f"Active Study     : {active_case.loc_name}")

    except Exception as err:

        print("ERROR")
        print(err)

    print("=" * 60)


# ----------------------------------------------------------
# Entry Point
# ----------------------------------------------------------

if __name__ == "__main__":
    main()