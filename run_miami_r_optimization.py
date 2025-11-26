#!/usr/bin/env python3
'''
Run the Miami-Dade county R code optimization (minimize number of new infections in last 5 years of simulation)
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
    parser.add_argument('-d', '--input_demographics_csv', required=True, type=str, help="Input abm_hiv-HRSA_SD Demographics CSV File")
    parser.add_argument('--scipy_minimize_method', required=False, type=str, default='Powell', help="SciPy Minimize Optimization Method (options: %s)" % ', '.join(sorted(SCIPY_MINIMIZE_METHODS)))
    parser.add_argument('--zip_output', action='store_true', help="Zip Output Files")
    parser.add_argument('--path_abm_hiv_commandline', required=False, type=str, default=DEFAULT_PATH_ABM_HIV_COMMANDLINE, help="Path to abm_hiv-HRSA_SD/abm_hiv_commandline.R")
    parser.add_argument('--path_abm_hiv_modules', required=False, type=str, default=DEFAULT_PATH_ABM_HIV_MODULES, help="Path to abm_hiv-HRSA_SD/modules")
    args = parser.parse_args()

    # process args
    args.output = Path(args.output)
    args.input_xlsx = Path(args.input_xlsx).expanduser().resolve()
    args.input_demographics_csv = Path(args.input_demographics_csv).expanduser().resolve()
    args.scipy_minimize_method = args.scipy_minimize_method.strip()
    args.path_abm_hiv_commandline = Path(args.path_abm_hiv_commandline).expanduser().resolve()
    args.path_abm_hiv_modules = Path(args.path_abm_hiv_modules).expanduser().resolve()

    # check args
    if args.output.exists():
        raise ValueError("Output exists: %s" % args.output)
    for p in [args.input_xlsx, args.input_demographics_csv, args.path_abm_hiv_commandline]:
        if not p.is_file():
            raise ValueError("File not found: %s" % p)
    for p in [args.path_abm_hiv_modules]:
        if not p.is_dir():
            raise ValueError("Directory not found: %s" % p)
    return args

# run ABM R code
def run_abm_hiv_r(out_path, abm_hiv_params_xlsx, abm_hiv_sd_demographics_csv, abm_hiv_trans_start=0.25, abm_hiv_trans_end=0.5, abm_hiv_trans_time=25, path_abm_hiv_commandline=DEFAULT_PATH_ABM_HIV_COMMANDLINE, path_abm_hiv_modules=DEFAULT_PATH_ABM_HIV_MODULES):
    out_path.mkdir(parents=True, exist_ok=True)
    command = ['Rscript', path_abm_hiv_commandline, path_abm_hiv_modules, abm_hiv_params_xlsx, abm_hiv_sd_demographics_csv, abm_hiv_trans_start, abm_hiv_trans_end, abm_hiv_trans_time]
    command = [str(x) for x in command] # convert to str
    print_log("Running abm_hiv-HRSA_SD Command: %s" % ' '.join(command))
    log_f = open(out_path / DEFAULT_FN_ABM_HIV_LOG, 'wt')
    log_f.write("=== ABM STDERR ===\n"); log_f.flush()
    abm_out = check_output(command, stderr=log_f).decode(); log_f.flush()
    log_f.write("\n\n=== ABM STDOUT ===\n"); log_f.write(abm_out); log_f.close()

# main execution
if __name__ == "__main__":
    # set things up
    args = parse_args()
    args.output.mkdir(parents=True)
    LOGFILE = open(args.output / DEFAULT_FN_LOG, 'wt')
    print_log("===== RUN INFORMATION =====")
    print_log("Calibration command: %s" % ' '.join(argv))
    print_log("Input abm_hiv-HRSA_SD Parameter XLSX: %s" % args.input_xlsx)
    print_log("Input abm_hiv-HRSA_SD Demographics CSV: %s" % args.input_demographics_csv)

    # copy input files to output directory
    input_copy_dir = args.output / 'inputs'
    input_copy_dir.mkdir()
    print_log("Copying input files to: %s" % input_copy_dir)
    for p in [args.input_xlsx, args.input_demographics_csv]:
        copy(p, input_copy_dir / p.name)
    print_log()

    # run calibration
    print_log("===== CALIBRATION =====")
    print_log("SciPy Optimize Calibration Method: %s" % args.scipy_minimize_method)
    print_log("Running calibration...")
    run_abm_hiv_r(args.output / 'test_run', args.input_xlsx, args.input_demographics_csv) # TODO DELETE AND MOVE TO WITHIN CALIBRATION
    pass # TODO RUN CALIBRATION
    if args.zip_output:
        print_log("Zipping output...")
        make_archive(args.output, 'zip', args.output)
        rmtree(args.output)
