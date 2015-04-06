from gurobipy import *
import argparse
import os
import sys
import tempfile

parser = argparse.ArgumentParser()
parser.add_argument("params", help="Additional Gurobi parameters to set", nargs="*")

args = parser.parse_args()
params_list = args.params

if os.name == 'nt':
    isWindows = True
else:
    isWindows = False

tFile = tempfile.gettempdir()
# On Mac, Excel doesn't write to the root of the temp folder, but to the
# 'TemporaryItems' folder within.
if not isWindows:
    tFile = os.path.join(tFile, 'TemporaryItems')

m = Model ('myModel')
m = read(os.path.join(tFile, 'model.lp'))
path = os.path.join(tFile, 'modelsolution.sol')

for param in params_list:
    name, value = param.split('=')
    try:
        m.setParam(name, float(value))
    except GurobiError as e:
        with open(path,'w') as File:
            File.write('Gurobi Error: %s' % e.message)
        sys.exit()

# Catch any GurobiError that occurs when solving
try:
    m.optimize()
except GurobiError as e:
    with open(path,'w') as File:
        File.write('Gurobi Error: %s' % e.message)
    sys.exit()

with open(path,'w') as File:
    File.write(str(m.status)+ '\n')
    if m.status != 3 and m.status != 4 and m.status != 5:
		m.write (path)
		Vars = m.getVars()
		Cons = m.getConstrs()
		with open (os.path.join(tFile, 'sensitivityData.sol'),'w') as destFile:
			for i in range(len(Vars)):
				try:
					destFile.write(str(Vars[i].rc)+','+str(Vars[i].SAObjLow)+','+str(Vars[i].SAObjUp)+'\n')
				except: pass
			for j in range(len(Cons)):
				try:
					destFile.write(str(Cons[j].pi)+',' + str(Cons[j].rhs)+','+ str(Cons[j].slack)+ ','  + str(Cons[j].SARHSLow) + ',' + str(Cons[j].SARHSUp) +'\n')
				except: pass
