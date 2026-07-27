"""
============================================================
04_activate_project.py

PowerFactory Python Toolkit

Purpose:
    Activate a PowerFactory project and verify that
    the activation was successful.

Author:
    Mazahir Hossein Learning Series

PowerFactory:
    2024

Python:
    3.11

Version:
    1.0
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
    """Connect to PowerFactory."""

    app = powerfactory.GetApplicationExt()

    if app is None:
        raise RuntimeError("Unable to connect to PowerFactory.")

    return app


# ----------------------------------------------------------
# Main
# ----------------------------------------------------------

def main():

    print("=" * 60)
    print("POWERFACTORY PROJECT ACTIVATION")
    print("=" * 60)

    try:

        app = connect_pf()

        print(f"Project Requested : {PROJECT_NAME}")
        print()

        # Activate project
        result = app.ActivateProject(PROJECT_NAME)

        if result != 0:
            print("Status           : FAILED")
            print(f"Return Code      : {result}")
            return

        # Verify activation
        project = app.GetActiveProject()

        if project is None:
            print("Status           : FAILED")
            print("Reason           : No active project found.")
            return

        print("Status           : SUCCESS")
        print(f"Active Project   : {project.loc_name}")

    except Exception as err:

        print("Status           : FAILED")
        print(err)

    print("=" * 60)


# ----------------------------------------------------------
# Entry Point
# ----------------------------------------------------------

if __name__ == "__main__":
    main()