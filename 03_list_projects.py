"""
============================================================
03_list_projects.py

PowerFactory Python Toolkit

Purpose:
    List all projects available for the current user.

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
    print("POWERFACTORY PROJECT LIST")
    print("=" * 60)

    try:

        app = connect_pf()

        user = app.GetCurrentUser()

        if user is None:
            print("Unable to retrieve current user.")
            return

        projects = user.GetContents("*.IntPrj", 1)

        if not projects:
            print("No projects found.")
            return

        print(f"Current User : {user.loc_name}")
        print()
        print("Available Projects")
        print("-" * 60)

        for index, project in enumerate(projects, start=1):

            print(f"{index:2d}. {project.loc_name}")

        print("-" * 60)
        print(f"Total Projects : {len(projects)}")

    except Exception as err:

        print("ERROR")
        print(err)

    print("=" * 60)


# ----------------------------------------------------------
# Entry Point
# ----------------------------------------------------------

if __name__ == "__main__":
    main()