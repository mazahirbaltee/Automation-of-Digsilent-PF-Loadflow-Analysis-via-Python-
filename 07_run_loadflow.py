"""
============================================================
07_run_loadflow.py

PowerFactory Python Toolkit

Purpose:
    Complete engineering load flow report.

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
    print("Unable to import PowerFactory.")
    raise SystemExit()


# ----------------------------------------------------------
# Functions
# ----------------------------------------------------------

def connect_pf():

    app = powerfactory.GetApplicationExt()

    if app is None:
        raise RuntimeError("Unable to connect.")

    return app


def activate_project(app):

    rc = app.ActivateProject(PROJECT_NAME)

    if rc != 0:
        raise RuntimeError("Project activation failed.")

    return app.GetActiveProject()


def activate_study_case(app):

    study_folder = app.GetProjectFolder("study")

    study_cases = study_folder.GetContents("*.IntCase", 1)

    for sc in study_cases:

        if sc.loc_name == STUDY_CASE_NAME:

            rc = sc.Activate()

            if rc != 1:
                raise RuntimeError("Study Case activation failed.")

            return sc

    raise RuntimeError("Study Case not found.")


def print_network_inventory(app):

    print("\n" + "=" * 60)
    print("NETWORK INVENTORY")
    print("=" * 60)

    inventory = [

        ("Buses (ElmTerm)", "*.ElmTerm"),
        ("Lines (ElmLne)", "*.ElmLne"),
        ("2-Winding Transformers (ElmTr2)", "*.ElmTr2"),
        ("3-Winding Transformers (ElmTr3)", "*.ElmTr3"),
        ("Loads (ElmLod)", "*.ElmLod"),
        ("Static Generators (ElmGenstat)", "*.ElmGenstat"),
        ("Synchronous Generators (ElmSym)", "*.ElmSym"),
        ("External Grid (ElmVac)", "*.ElmVac"),
        ("Shunts (ElmShnt)", "*.ElmShnt"),

    ]

    counts = {}

    for title, pattern in inventory:

        objs = app.GetCalcRelevantObjects(pattern)

        if objs is None:
            objs = []

        counts[title] = len(objs)

        print(f"{title:<35}{len(objs):>5}")

    return counts


def run_loadflow(app):

    print("\n" + "=" * 60)
    print("LOAD FLOW")
    print("=" * 60)

    ldf = app.GetFromStudyCase("ComLdf")

    if ldf is None:
        raise RuntimeError("Load Flow object not found.")

    rc = ldf.Execute()

    print(f"Status       : {'SUCCESS' if rc == 0 else 'FAILED'}")
    print(f"Return Code  : {rc}")

    if rc != 0:
        raise RuntimeError("Load Flow failed.")


def print_bus_voltages(app):

    print("\n" + "=" * 60)
    print("BUS VOLTAGES")
    print("=" * 60)

    print(f"{'Bus Name':<25}{'Voltage (pu)':>15}")

    buses = app.GetCalcRelevantObjects("*.ElmTerm")

    vmax = -1
    vmin = 99

    for bus in buses:

        try:
            u = bus.GetAttribute("m:u")
        except:
            continue

        print(f"{bus.loc_name:<25}{u:>15.4f}")

        vmax = max(vmax, u)
        vmin = min(vmin, u)

    return vmin, vmax


def print_transformers(app):

    print("\n" + "=" * 60)
    print("TRANSFORMER RESULTS")
    print("=" * 60)

    trs = app.GetCalcRelevantObjects("*.ElmTr2")

    if len(trs) == 0:

        print("No transformers found.")

        return 0

    max_loading = 0

    print(f"{'Name':<28}{'Loading%':>12}{'P(MW)':>12}{'Q(MVAr)':>12}")

    for tr in trs:

        loading = tr.GetAttribute("c:loading")
        p = tr.GetAttribute("m:P:bushv")
        q = tr.GetAttribute("m:Q:bushv")

        max_loading = max(max_loading, loading)

        print(
            f"{tr.loc_name:<28}"
            f"{loading:>12.2f}"
            f"{p:>12.2f}"
            f"{q:>12.2f}"
        )

    return max_loading


def print_loads(app):

    print("\n" + "=" * 60)
    print("LOAD INPUTS")
    print("=" * 60)

    loads = app.GetCalcRelevantObjects("*.ElmLod")

    if len(loads) == 0:

        print("No loads found.")
        return

    print(f"{'Load Name':<20}{'P(MW)':>12}{'Q(MVAr)':>12}")

    total_p = 0
    total_q = 0

    for load in loads:

        p = load.plini
        q = load.qlini

        total_p += p
        total_q += q

        print(f"{load.loc_name:<20}{p:>12.2f}{q:>12.2f}")

    print("-" * 44)
    print(f"{'TOTAL':<20}{total_p:>12.2f}{total_q:>12.2f}")

    return total_p, total_q


def print_summary(counts, vmin, vmax, max_loading, total_p, total_q):

    print("\n" + "=" * 60)
    print("SYSTEM SUMMARY")
    print("=" * 60)

    print(f"Total Buses              : {counts['Buses (ElmTerm)']}")
    print(f"Total Loads              : {counts['Loads (ElmLod)']}")
    print(f"Total Transformers       : {counts['2-Winding Transformers (ElmTr2)']}")
    print()

    print(f"Total Load (MW)          : {total_p:.2f}")
    print(f"Total Reactive (MVAr)    : {total_q:.2f}")
    print()

    print(f"Maximum Voltage (pu)     : {vmax:.4f}")
    print(f"Minimum Voltage (pu)     : {vmin:.4f}")
    print()

    print(f"Maximum Transformer Load : {max_loading:.2f} %")

    print("=" * 60)


# ----------------------------------------------------------
# Main
# ----------------------------------------------------------

def main():

    app = connect_pf()

    activate_project(app)

    activate_study_case(app)

    counts = print_network_inventory(app)

    run_loadflow(app)

    vmin, vmax = print_bus_voltages(app)

    max_loading = print_transformers(app)

    total_p, total_q = print_loads(app)

    print_summary(
        counts,
        vmin,
        vmax,
        max_loading,
        total_p,
        total_q
    )


# ----------------------------------------------------------
# Entry Point
# ----------------------------------------------------------

if __name__ == "__main__":
    main()