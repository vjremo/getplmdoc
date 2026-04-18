#!/usr/bin/env python3
"""Combined RTM runner: LCS, JSP, CSP, and Techpack."""

import argparse
import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent
SCRIPTS = {
    "lcs": BASE_DIR / "ssp-rtm-sync" / "ssp.py",
    "jsp": BASE_DIR / "jsp-rtm-sync" / "scripts" / "jsp.py",
    "csp": BASE_DIR / "csp-rtm-sync" / "scripts" / "csp.py",
    "techpack": BASE_DIR / "techpack-rtm-sync" / "techpack.py",
}


def section(title):
    print(f"\n{'='*60}")
    print(f"  {title}")
    print(f"{'='*60}")


def run(cmd, cwd=BASE_DIR):
    result = subprocess.run(cmd, cwd=cwd)
    if result.returncode != 0:
        print(f"[ERROR] Command exited with code {result.returncode}", file=sys.stderr)
    return result.returncode


def run_lcs(args):
    section("LCS RTM")
    cmd = [sys.executable, str(SCRIPTS["lcs"]), "-o", args.lcs_output]
    if args.lcs_input:
        cmd += args.lcs_input  # explicit files; otherwise convert.py globs cwd
    if args.lcs_template:
        cmd += ["-t", args.lcs_template]
    return run(cmd)


def run_jsp(args):
    section("JSP RTM")
    return run([sys.executable, str(SCRIPTS["jsp"])], cwd=args.work_dir)


def run_csp(args):
    section("CSP RTM")
    cmd = [
        sys.executable, str(SCRIPTS["csp"]),
        "--properties", args.csp_properties,
        "--rtm", args.csp_rtm,
    ]
    return run(cmd)


def run_techpack(args):
    section("Techpack RTM")
    return run([sys.executable, str(SCRIPTS["techpack"])], cwd=args.work_dir)


def main():
    parser = argparse.ArgumentParser(
        description="Run all RTM update scripts: LCS, JSP, CSP, and Techpack."
    )

    parser.add_argument(
        "--only", nargs="+", choices=["lcs", "jsp", "csp", "techpack"],
        metavar="SCRIPT",
        help="Run only the specified scripts (default: all). Choices: lcs, jsp, csp, techpack",
    )
    parser.add_argument(
        "--work-dir", default=str(BASE_DIR),
        help="Working directory for JSP and Techpack scripts (default: project root)",
    )

    # LCS options
    lcs_group = parser.add_argument_group("LCS RTM (ssp-rtm-sync/ssp.py)")
    lcs_group.add_argument("--lcs-input", nargs="*", default=[],
                           metavar="FILE",
                           help="LCS .properties input file(s). Defaults to all *.properties in properties/ folder.")
    lcs_group.add_argument("--lcs-output", default="SSP_RTM.xlsx",
                           metavar="FILE", help="LCS RTM output .xlsx file (default: SSP_RTM.xlsx)")
    lcs_group.add_argument("--lcs-template", default=None,
                           metavar="FILE", help="Optional existing SSP_RTM.xlsx to use as template")

    # CSP options
    csp_group = parser.add_argument_group("CSP RTM (csp-rtm-sync/scripts/csp.py)")
    csp_group.add_argument(
        "--csp-properties",
        default="properties/custom.clientSidePluginManagerMappings.properties",
        metavar="FILE", help="CSP .properties input file",
    )
    csp_group.add_argument("--csp-rtm", default="CSP_RTM.xlsx",
                           metavar="FILE", help="CSP RTM .xlsx file (default: CSP_RTM.xlsx)")

    args = parser.parse_args()
    args.work_dir = Path(args.work_dir).resolve()

    scripts_to_run = args.only or ["lcs", "jsp", "csp", "techpack"]
    runners = {"lcs": run_lcs, "jsp": run_jsp, "csp": run_csp, "techpack": run_techpack}

    errors = 0
    for name in scripts_to_run:
        errors += runners[name](args) != 0

    print(f"\n{'='*60}")
    if errors:
        print(f"  Done with {errors} error(s).")
        sys.exit(1)
    else:
        print("  All RTM scripts completed successfully.")


if __name__ == "__main__":
    main()
