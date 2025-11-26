#!/usr/bin/env python3
'''
Run the Miami-Dade county R code optimization (minimize number of new infections in last 5 years of simulation)
'''

# imports
from datetime import datetime
from openpyxl import load_workbook
from os import cpu_count
from pathlib import Path
from random import randint
from scipy.optimize import LinearConstraint, minimize
from shutil import copy, make_archive, rmtree
from statistics import mean
from subprocess import check_output, run
from sys import argv, stdout
import argparse

# constants
LOGFILE = None
MAX_RNG_SEED = 99999
SCIPY_MINIMIZE_METHODS = {'Nelder-Mead', 'L-BFGS-B', 'TNC', 'SLSQP', 'Powell', 'trust-constr', 'COBYLA', 'COBYQA'}
OPTIMIZATION_MODES = {'geo', 'risk', 'race'}
SCORE_FUNCTIONS = {'mean': mean, 'max': max, 'min': min}

# defaults
DEFAULT_NUM_REPS_PER_SCORE = 5
DEFAULT_MAX_NUM_THREADS = cpu_count()
DEFAULT_SCORE_FUNCTION = 'mean'
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
    parser.add_argument('-m', '--optimization_mode', required=True, type=str, help="Optimization Mode (options: %s)" % ', '.join(sorted(OPTIMIZATION_MODES)))
    parser.add_argument('-n', '--num_reps_per_score', required=False, type=int, default=DEFAULT_NUM_REPS_PER_SCORE, help="Number of Simulation Replicates per Score")
    parser.add_argument('-t', '--max_num_threads', required=False, type=int, default=DEFAULT_MAX_NUM_THREADS, help="Max Number of Threads to Use")
    parser.add_argument('--score_function', required=False, type=str, default=DEFAULT_SCORE_FUNCTION, help="Function to Calculate Overall Score from Replicates (options: %s)" % ', '.join(sorted(SCORE_FUNCTIONS.keys())))
    parser.add_argument('--scipy_minimize_method', required=False, type=str, default='Powell', help="SciPy Minimize Optimization Method (options: %s)" % ', '.join(sorted(SCIPY_MINIMIZE_METHODS)))
    parser.add_argument('--zip_output', action='store_true', help="Zip Output Files")
    parser.add_argument('--path_abm_hiv_commandline', required=False, type=str, default=DEFAULT_PATH_ABM_HIV_COMMANDLINE, help="Path to abm_hiv-HRSA_SD/abm_hiv_commandline.R")
    parser.add_argument('--path_abm_hiv_modules', required=False, type=str, default=DEFAULT_PATH_ABM_HIV_MODULES, help="Path to abm_hiv-HRSA_SD/modules")
    args = parser.parse_args()

    # process args
    args.output = Path(args.output)
    args.input_xlsx = Path(args.input_xlsx).expanduser().resolve()
    args.input_demographics_csv = Path(args.input_demographics_csv).expanduser().resolve()
    args.optimization_mode = args.optimization_mode.strip().lower()
    args.score_function = args.score_function.strip().lower()
    args.scipy_minimize_method = args.scipy_minimize_method.strip()
    args.path_abm_hiv_commandline = Path(args.path_abm_hiv_commandline).expanduser().resolve()
    args.path_abm_hiv_modules = Path(args.path_abm_hiv_modules).expanduser().resolve()

    # check args
    if args.optimization_mode not in OPTIMIZATION_MODES:
        raise ValueError("Invalid optimization mode (%s). Options: %s" % (args.optimization_mode, ', '.join(sorted(OPTIMIZATION_MODES))))
    if args.score_function not in SCORE_FUNCTIONS:
        raise ValueError("Invalid score function (%s). Options: %s" % (args.score_function, ', '.join(sorted(SCORE_FUNCTIONS.keys()))))
    if args.scipy_minimize_method not in SCIPY_MINIMIZE_METHODS:
        raise ValueError("Invalid SciPy minimize method (%s). Options: %s" % (args.scipy_minimize_method, ', '.join(sorted(SCIPY_MINIMIZE_METHODS))))
    if args.num_reps_per_score < 1:
        raise ValueError("Number of simulation replicates per score must be positive: %d" % args.num_reps_per_score)
    if args.max_num_threads < 1:
        raise ValueError("Maximum number of threads must be positive: %d" % args.max_num_threads)
    if args.output.exists():
        raise ValueError("Output exists: %s" % args.output)
    for p in [args.input_xlsx, args.input_demographics_csv, args.path_abm_hiv_commandline]:
        if not p.is_file():
            raise ValueError("File not found: %s" % p)
    for p in [args.path_abm_hiv_modules]:
        if not p.is_dir():
            raise ValueError("Directory not found: %s" % p)
    return args

# perform R code optimization
def run_optimization(
    mode, out_path, abm_hiv_params_xlsx, abm_hiv_sd_demographics_csv,
    abm_hiv_trans_start=0.25, abm_hiv_trans_end=0.5, abm_hiv_trans_time=25,
    num_reps_per_score=DEFAULT_NUM_REPS_PER_SCORE, max_num_threads=DEFAULT_MAX_NUM_THREADS, score_func=SCORE_FUNCTIONS[DEFAULT_SCORE_FUNCTION],
    path_abm_hiv_commandline=DEFAULT_PATH_ABM_HIV_COMMANDLINE, path_abm_hiv_modules=DEFAULT_PATH_ABM_HIV_MODULES
    ):
    # prep optimization parameters
    if mode == 'geo':
        top_row = 12
        num_cells = 13
    elif mode == 'risk':
        top_row = 27
        num_cells = 4
    elif mode == 'race':
        top_row = 33
        num_cells = 3
    else:
        raise ValueError("Invalid optimization mode (%s). Options: %s" % (mode, ', '.join(sorted(OPTIMIZATION_MODES))))
    x0 = [1./num_cells] * num_cells
    bounds = [(0,1)] * num_cells
    linear_constraint = LinearConstraint([[1]*num_cells], 1, 1)

    # nested optimization function
    iter_num = 0
    def opt_fun(x):
        # prep ABM R code run
        nonlocal iter_num; iter_num += 1
        print_log("Preparing ABM iteration %d..." % iter_num)
        curr_out_path = out_path / ('optimization.iteration.%s' % str(iter_num).zfill(3))
        curr_out_path.mkdir(parents=True)
        rep_nums = [str(i).zfill(3) for i in range(1, num_reps_per_score+1)]

        # prep ABM XLSX files
        rng_seed_base = randint(0, MAX_RNG_SEED)
        for rep_num in rep_nums:
            data_xlsx_copy_path = curr_out_path / ('rep.%s.data.xlsx' % rep_num)
            wb = load_workbook(abm_hiv_params_xlsx, data_only=True) # load original XLSX
            wb['High Level Pop + Sim Features']['B4'] = rng_seed_base + int(rep_num) # override RNG seed
            for i in range(num_cells):
                wb['Testing']['M%d' % (top_row+i)] = x[i]
            for ws in wb.worksheets:
                for row in ws:
                    for cell in row:
                        if isinstance(cell.value, str):
                            cell.value = cell.value.replace('\n', '\r\n')
            wb.save(data_xlsx_copy_path); wb.close()

        # run ABM R code
        command = [ # base command
            'parallel', '--jobs', str(max_num_threads),
            'Rscript', str(path_abm_hiv_commandline), str(path_abm_hiv_modules), str(curr_out_path / 'rep.{}.data.xlsx'), str(abm_hiv_sd_demographics_csv), str(abm_hiv_trans_start), str(abm_hiv_trans_end), str(abm_hiv_trans_time),
            '>', str(curr_out_path / 'rep.{}.log.txt'), '2>&1',
        ]
        command += [':::'] + rep_nums
        print_log("Running ABM iteration %d: %s" % (iter_num, ' '.join("'%s'" % c for c in command)))
        run(command)

        # compute score from outputs
        transmission_counts = list()
        for rep_num in rep_nums:
            with open(curr_out_path / ('rep.%s.log.txt' % rep_num), 'rt') as f:
                transmissions = [[x.strip() for x in l.strip().split()] for l in f.read().split('[1] "Transmission tree..."')[1].split('[1] "Sequence sample times..."').strip().splitlines()]
            end_month = max(t for u, v, t in transmissions)
            month_threshold = end_month - (12*score_num_years)
            transmission_counts.append(sum(1 for u, v, t in transmissions if t > month_threshold))
        return score_func(transmission_counts)
    return minimize(opt_fun, x0, constraints=[linear_constraint])

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
    print_log("SciPy Optimize Minimize Method: %s" % args.scipy_minimize_method)
    print_log("Running optimization...")
    results = run_optimization(
        args.optimization_mode, args.output / 'simulations', args.input_xlsx, args.input_demographics_csv,
        num_reps_per_score=args.num_reps_per_score, max_num_threads=args.max_num_threads, score_func=SCORE_FUNCTIONS[args.score_function],
        path_abm_hiv_commandline=args.path_abm_hiv_commandline, path_abm_hiv_modules=args.path_abm_hiv_modules,
    )
    print_log("Best parameters: %s" % str(results.x))
    if args.zip_output:
        print_log("Zipping output...")
        make_archive(args.output, 'zip', args.output)
        rmtree(args.output)
