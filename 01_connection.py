"""
============================================================
01_connection.py
PowerFactory Python Toolkit

Purpose:
    Verify that Python can successfully connect to
    DIgSILENT PowerFactory.

Author:
    Mazahir Hossein Learning Series
Version:
    1.0
============================================================
"""

import sys

# ------------------------------------------------------------------
# PowerFactory Python API Path
# ------------------------------------------------------------------
PF_PATH = r"C:\Program Files\DIgSILENT\PowerFactory 2024\Python\3.11"

if PF_PATH not in sys.path:
    sys.path.append(PF_PATH)

try:
    import powerfactory
except ImportError:
    print("ERROR : Unable to import the PowerFactory Python API.")
    print("Check PF_PATH.")
    raise SystemExit()


def connect_pf():
    """
    Connect to PowerFactory Engine Mode.

    Returns
    -------
    app : PowerFactory Application object
    """

    app = powerfactory.GetApplicationExt()

    if app is None:
        raise RuntimeError("Unable to connect to PowerFactory.")

    return app


def main():

    print("=" * 60)
    print("POWERFACTORY CONNECTION TEST")
    print("=" * 60)

    try:

        app = connect_pf()

        print("Status : Connected Successfully")

        print("Application :", app)

        print("=" * 60)

    except Exception as err:

        print("Status : Connection Failed")

        print(err)

        print("=" * 60)


if __name__ == "__main__":
    main()