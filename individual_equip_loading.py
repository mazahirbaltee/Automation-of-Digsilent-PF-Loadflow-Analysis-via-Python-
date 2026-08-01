import sys
import time
import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from pathlib import Path

# ============================================================
# USER SETTINGS
# ============================================================
N_STEPS         = 10
SCALE_START     = 1.0          # 100 
SCALE_END       = 3.5         # 350 %
PAUSE_SECONDS   = 5

# Correct PowerFactory Python path for 2024 + Python 3.11
PF_PYTHON_PATH  = r"C:\Program Files\DIgSILENT\PowerFactory 2024\Python\3.11"

PROJECT_NAME    = "G99_RPC2"       # ← change this
STUDY_CASE_NAME = "Study Case"            # ← change this

# Output folder
OUTPUT_DIR = Path(r"D:\DigSilent_PowerFactory_Lessons\Python_Integ\Grook_PF")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
EXCEL_FILE = OUTPUT_DIR / "individual_loadings_results.xlsx"
# ============================================================

sys.path.append(PF_PYTHON_PATH)
import powerfactory as pf

app = pf.GetApplication()
if app is None:
    raise Exception("Cannot connect to PowerFactory. Open PowerFactory + project first.")

app.Show()
app.ClearOutputWindow()
app.PrintPlain("=== Individual Loading Script (Seaborn + Excel) ===")

# ----------------------------------------------------------------------
# Activate project / study case
# ----------------------------------------------------------------------
if PROJECT_NAME:
    app.ActivateProject(PROJECT_NAME)

if STUDY_CASE_NAME:
    folder = app.GetProjectFolder("study")
    cases = folder.GetContents(STUDY_CASE_NAME + ".IntCase")
    if cases:
        cases[0].Activate()
        app.PrintPlain(f"Study case activated: {STUDY_CASE_NAME}")

# ----------------------------------------------------------------------
# Collect all objects
# ----------------------------------------------------------------------
loads        = app.GetCalcRelevantObjects("*.ElmLod")
lines        = app.GetCalcRelevantObjects("*.ElmLne")
transformers = app.GetCalcRelevantObjects("*.ElmTr2") + app.GetCalcRelevantObjects("*.ElmTr3")
generators   = app.GetCalcRelevantObjects("*.ElmSym") + app.GetCalcRelevantObjects("*.ElmGenstat")

if not loads:
    raise Exception("No loads found!")

original_loads = {ld: (ld.plini, ld.qlini) for ld in loads}

print(f"Loads: {len(loads)} | Lines: {len(lines)} | "
      f"Transformers: {len(transformers)} | Generators: {len(generators)}")
app.PrintPlain(f"Loads:{len(loads)}  Lines:{len(lines)}  Trafos:{len(transformers)}  Gens:{len(generators)}")

def get_loading(obj):
    try:
        val = getattr(obj, "c:loading")
        return float(val) if val is not None else np.nan
    except Exception:
        try:
            return float(obj.GetAttribute("c:loading"))
        except Exception:
            return np.nan

def restore_original():
    for ld, (p, q) in original_loads.items():
        ld.plini = p
        ld.qlini = q
    app.PrintPlain(">>> Original load values restored <<<")
    print(">>> Original load values restored <<<")

# ----------------------------------------------------------------------
# Load-flow command
# ----------------------------------------------------------------------
ldf = app.GetFromStudyCase("ComLdf")
ldf.iopt_net = 0

# ----------------------------------------------------------------------
# BASE-CASE check
# ----------------------------------------------------------------------
app.PrintPlain("--- BASE-CASE (100 %) ---")
ierr = ldf.Execute()
if ierr != 0:
    print(f"BASE-CASE failed (error {ierr}). Fix network first!")
    app.PrintError(f"BASE-CASE failed (error {ierr})")
    sys.exit(1)
print("BASE-CASE OK")

# ----------------------------------------------------------------------
# Main loop
# ----------------------------------------------------------------------
scale_factors = np.linspace(SCALE_START, SCALE_END, N_STEPS)
records = []          # long-format: one row per element per step

try:
    for step, factor in enumerate(scale_factors, 1):
        # Scale ALL loads to the same percentage
        for ld in loads:
            ld.plini = original_loads[ld][0] * factor
            ld.qlini = original_loads[ld][1] * factor

        ierr = ldf.Execute()
        if ierr != 0:
            msg = (f"Load-flow FAILED at step {step}/{N_STEPS} "
                   f"(Scale = {factor*100:.1f} %)  error={ierr}")
            print(msg)
            app.PrintError(msg)
            restore_original()
            sys.exit(1)

        scale_pct = round(factor * 100, 1)

        # ----- Record every individual element -----
        for ld in loads:
            records.append({
                "Step": step,
                "Scale [%]": scale_pct,
                "Type": "Load",
                "Name": ld.loc_name,
                "Loading [%]": scale_pct,          # the applied scale
                "P [MW]": ld.plini,
                "Q [Mvar]": ld.qlini
            })

        for ln in lines:
            records.append({
                "Step": step,
                "Scale [%]": scale_pct,
                "Type": "Line",
                "Name": ln.loc_name,
                "Loading [%]": get_loading(ln),
                "P [MW]": np.nan,
                "Q [Mvar]": np.nan
            })

        for tr in transformers:
            records.append({
                "Step": step,
                "Scale [%]": scale_pct,
                "Type": "Transformer",
                "Name": tr.loc_name,
                "Loading [%]": get_loading(tr),
                "P [MW]": np.nan,
                "Q [Mvar]": np.nan
            })

        for gen in generators:
            records.append({
                "Step": step,
                "Scale [%]": scale_pct,
                "Type": "Generator",
                "Name": gen.loc_name,
                "Loading [%]": get_loading(gen),
                "P [MW]": np.nan,
                "Q [Mvar]": np.nan
            })

        # Console summary (max only)
        def max_of(typ):
            vals = [r["Loading [%]"] for r in records
                    if r["Step"] == step and r["Type"] == typ and not np.isnan(r["Loading [%]"])]
            return max(vals) if vals else np.nan

        print(f"Step {step:02d}/{N_STEPS} | {scale_pct:6.1f}% | "
              f"MaxLine:{max_of('Line'):7.1f}  MaxTrafo:{max_of('Transformer'):7.1f}  "
              f"MaxGen:{max_of('Generator'):7.1f}  [OK]")

        time.sleep(PAUSE_SECONDS)

except Exception as e:
    print(f"Unexpected error: {e}")
    restore_original()
    raise

# ----------------------------------------------------------------------
# Success → restore + save Excel + plots
# ----------------------------------------------------------------------
restore_original()

df = pd.DataFrame(records)

# ---------- Save to Excel (multiple sheets) ----------
with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="All_Data", index=False)

    # Separate sheets for convenience
    for typ in ["Load", "Line", "Transformer", "Generator"]:
        df[df["Type"] == typ].to_excel(writer, sheet_name=typ, index=False)

print(f"\nExcel file saved to:\n{EXCEL_FILE}")
app.PrintPlain(f"Excel saved: {EXCEL_FILE}")

# ---------- Seaborn plots (three separate figures) ----------
sns.set_theme(style="whitegrid", font_scale=1.1)

def plot_individual(data, title, filename):
    if data.empty:
        print(f"No data for {title}")
        return
    plt.figure(figsize=(12, 6))
    ax = sns.lineplot(
        data=data,
        x="Scale [%]",
        y="Loading [%]",
        hue="Name",
        marker="o",
        linewidth=1.8
    )
    ax.axhline(100, color="red", linestyle="--", alpha=0.6, label="100 % limit")
    ax.set_title(title)
    ax.set_xlabel("Load Scaling [%]")
    ax.set_ylabel("Loading [%]")
    plt.legend(bbox_to_anchor=(1.05, 1), loc="upper left", fontsize=8)
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / filename, dpi=150, bbox_inches="tight")
    plt.show()

# 1. Lines
plot_individual(
    df[df["Type"] == "Line"],
    "Individual Line Loadings vs Load Scaling",
    "lines_loading.png"
)

# 2. Transformers
plot_individual(
    df[df["Type"] == "Transformer"],
    "Individual Transformer Loadings vs Load Scaling",
    "transformers_loading.png"
)

# 3. Generators
plot_individual(
    df[df["Type"] == "Generator"],
    "Individual Generator Loadings vs Load Scaling",
    "generators_loading.png"
)

print("\nAll three seaborn plots displayed and saved as PNG.")
app.PrintPlain("Script finished successfully.")
print("Script finished successfully.")