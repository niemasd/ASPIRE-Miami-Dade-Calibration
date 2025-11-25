#!/usr/bin/env python3
'''
Run the Miami-Dade county R code calibration
'''

# imports
from datetime import datetime
from openpyxl import load_workbook
from pathlib import Path
from shutil import copy, make_archive, rmtree
from subprocess import check_output
from sys import argv, stdout
import argparse

# constants
LOGFILE = None
MAX_RNG_SEED = 2147483647
SCIPY_MINIMIZE_METHODS = {'Nelder-Mead', 'L-BFGS-B', 'TNC', 'SLSQP', 'Powell', 'trust-constr', 'COBYLA', 'COBYQA'}

# defaults
DEFAULT_PATH_ABM_HIV_COMMANDLINE = "/usr/local/bin/abm_hiv-HRSA_SD/abm_hiv_commandline.R"
DEFAULT_PATH_ABM_HIV_MODULES = "/usr/local/bin/abm_hiv-HRSA_SD/modules"
DEFAULT_FN_ABM_HIV_LOG = "log_abm_hiv.txt"
DEFAULT_FN_LOG = "log_calibration.txt"

# check GNU parallel
if not check_output(['parallel', '-h']).decode().strip().startswith('Usage:'):
    raise RuntimeError("GNU `parallel` does not seem to be installed.\nInstall with e.g.: sudo apt-get install -y parallel")

# return the current time as a string
def get_time():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# print to the log (None implies stdout only)
def print_log(s='', end='\n'):
    tmp = "[%s] %s" % (get_time(), s)
    print(tmp, end=end); stdout.flush()
    if LOGFILE is not None:
        print(tmp, file=LOGFILE, end=end); LOGFILE.flush()

# parse user args
def parse_args():
    # parse args
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument('-o', '--output', required=True, type=str, help="Output Directory")
    parser.add_argument('-x', '--input_xlsx', required=True, type=str, help="Input abm_hiv-HRSA_SD Parameter XLSX File")
    parser.add_argument('--scipy_minimize_method', required=False, type=str, default='Powell', help="SciPy Minimize Optimization Method (options: %s)" % ', '.join(sorted(SCIPY_MINIMIZE_METHODS)))
    parser.add_argument('--zip_output', action='store_true', help="Zip Output Files")
    parser.add_argument('--path_abm_hiv_commandline', required=False, type=str, default=DEFAULT_PATH_ABM_HIV_COMMANDLINE, help="Path to abm_hiv-HRSA_SD/abm_hiv_commandline.R")
    parser.add_argument('--path_abm_hiv_modules', required=False, type=str, default=DEFAULT_PATH_ABM_HIV_MODULES, help="Path to abm_hiv-HRSA_SD/modules")
    args = parser.parse_args()

    # process args
    args.output = Path(args.output)
    args.input_xlsx = Path(args.input_xlsx).expanduser().absolute()
    args.scipy_minimize_method = args.scipy_minimize_method.strip()
    args.path_abm_hiv_commandline = Path(args.path_abm_hiv_commandline).expanduser().absolute()
    args.path_abm_hiv_modules = Path(args.path_abm_hiv_modules).expanduser().absolute()

    # check args
    if args.output.exists():
        raise ValueError("Output exists: %s" % args.output)
    for p in [args.input_xlsx, args.path_abm_hiv_commandline]:
        if not p.is_file():
            raise ValueError("File not found: %s" % p)
    for p in [args.path_abm_hiv_modules]:
        if not p.is_dir():
            raise ValueError("Directory not found: %s" % p)
    return args

# main execution
if __name__ == "__main__":
    # set things up
    args = parse_args()
    args.output.mkdir(parents=True)
    LOGFILE = open(args.output / DEFAULT_FN_LOG, 'wt')
    print_log("===== RUN INFORMATION =====")
    print_log("Calibration command: %s" % ' '.join(argv))
    print_log("Original abm_hiv-HRSA_SD Parameter XLSX: %s" % args.input_xlsx)
    copy(args.input_xlsx, args.output / 'input.xlsx')
    print_log()

    # run calibration
    print_log("===== CALIBRATION =====")
    print_log("SciPy Optimize Calibration Method: %s" % args.scipy_minimize_method)
    print_log("Running calibration...")
    pass # TODO RUN CALIBRATION
    if args.zip_output:
        print_log("Zipping output...")
        make_archive(args.output, 'zip', args.output)
        rmtree(args.output)
