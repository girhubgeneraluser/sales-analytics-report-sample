"""Refresh the whole report: clean -> analyze -> verify.

Sample project - synthetic data, not client work.

    python3 src/run_all.py                    # refresh from whatever is in data/raw/
    python3 src/run_all.py --with-sample-data # regenerate the synthetic raw data first
"""
import subprocess
import sys
import time
from pathlib import Path

SRC = Path(__file__).resolve().parent
steps = (["generate_data.py"] if "--with-sample-data" in sys.argv else []) + ["clean.py", "analyze.py", "verify.py"]
t0 = time.time()
for step in steps:
    print(f"\n=== {step} ===", flush=True)
    if step == "verify.py":
        res = subprocess.run([sys.executable, str(SRC / step), "--verbose"], capture_output=True, text=True)
        (SRC.parent / "deliverables" / "qa_report.txt").write_text(res.stdout)
        print("\n".join(l for l in res.stdout.splitlines() if not l.startswith("PASS")))
        rc = res.returncode
    else:
        rc = subprocess.call([sys.executable, str(SRC / step)])
    if rc != 0:
        print(f"\nSTOPPED: {step} exited with code {rc}")
        sys.exit(rc)
print(f"\nDone in {time.time() - t0:.1f}s. Outputs in deliverables/.")
