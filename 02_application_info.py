"""
============================================================
02_application_info.py

PowerFactory Python Toolkit

Purpose:
    Display information about the current
    PowerFactory application session.

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

    app = powerfactory.GetApplicationExt()

    if app is None:
        raise RuntimeError("Unable to connect to PowerFactory.")

    return app


# ----------------------------------------------------------
# Main
# ----------------------------------------------------------

def main():

    print("=" * 60)
    print("POWERFACTORY APPLICATION INFORMATION")
    print("=" * 60)

    try:

        app = connect_pf()

        print(f"Status          : Connected")
        print(f"Application     : {app}")

        # --------------------------------------------------
        # Current User
        # --------------------------------------------------

        try:

            user = app.GetCurrentUser()

            if user:
                print(f"Current User    : {user.loc_name}")
            else:
                print("Current User    : None")

        except Exception:
            print("Current User    : Unable to retrieve")

        # --------------------------------------------------
        # Active Project
        # --------------------------------------------------

        try:

            project = app.GetActiveProject()

            if project:
                print(f"Active Project  : {project.loc_name}")
            else:
                print("Active Project  : None")

        except Exception:
            print("Active Project  : Unable to retrieve")

        # --------------------------------------------------
        # Active Study Case
        # --------------------------------------------------

        try:

            studycase = app.GetActiveStudyCase()

            if studycase:
                print(f"Study Case      : {studycase.loc_name}")
            else:
                print("Study Case      : None")

        except Exception:
            print("Study Case      : Unable to retrieve")

    except Exception as err:

        print("Status          : Connection Failed")
        print(err)

    print("=" * 60)


# ----------------------------------------------------------
# Entry Point
# ----------------------------------------------------------

if __name__ == "__main__":
    main()