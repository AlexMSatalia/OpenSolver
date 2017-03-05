import argparse
import os
import sys
import tempfile

# Catch gurobi licensing errors
try:
  from gurobipy import *
except Exception as e:
  print('%s' % e.message)
  sys.exit(1)

parser = argparse.ArgumentParser()
parser.add_argument("modelfile", help="Path to the model.lp file")
parser.add_argument("mipstartfile", help="Path to any mipstart .mst file")
parser.add_argument("statusfile", help="Path to write the status file")
parser.add_argument("solutionfile", help="Path to write the solution .sol file")
parser.add_argument("sensitivityfile",
                    help="Path to write the sensitivity data")
parser.add_argument("params", help="Additional Gurobi parameters to set",
                    nargs="*")

args = parser.parse_args()
mod_path = args.modelfile
mst_path = args.mipstartfile
status_path = args.statusfile
sol_path = args.solutionfile
sense_path = args.sensitivityfile
params_list = args.params

m = Model('myModel')
m = read(mod_path)

if os.path.exists(mst_path):
    m.read(mst_path)

for param in params_list:
    name, value = param.split('=')
    try:
        m.setParam(name, float(value))
    except GurobiError as e:
        with open(status_path, 'w') as status_file:
            status_file.write('Gurobi Error: %s' % e.message)
        sys.exit()

# Catch any GurobiError that occurs when solving
try:
    m.optimize()
except GurobiError as e:
    with open(status_path, 'w') as status_file:
        status_file.write('Gurobi Error: %s' % e.message)
    sys.exit()

with open(status_path, 'w') as status_file:
    status_file.write(str(m.status) + '\n')

if not m.status in frozenset((3, 4, 5)):
    m.write(sol_path)
    vars = m.getVars()
    cons = m.getConstrs()
    with open(sense_path, 'w') as sense_file:
        try:
            vals = [map(str, m.getAttr(t, vars))
                    for t in ['RC', 'SAObjLow', 'SAObjUp']]
            for k in zip(*vals):
                sense_file.write(','.join(k) + '\n')
        except:
            pass
        try:
            vals = [map(str, m.getAttr(t, cons))
                    for t in ['Pi', 'RHS', 'Slack', 'SARHSLow', 'SARHSUp']]
            for k in zip(*vals):
                sense_file.write(','.join(k) + '\n')
        except:
            pass
