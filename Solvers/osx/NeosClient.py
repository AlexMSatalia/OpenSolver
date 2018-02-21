#!/usr/bin/env python

# Python script for contacting NEOS over XML-RPC
# NeosClient.py is included in NEOS download package:
# http://www.neos-server.org/neos/downloads.html
# Modifications by JWD

import sys
import xmlrpclib
import os

NEOS_HOST="www.neos-server.org"
NEOS_PORT=3333

numargs = {
  "ping":  3,
  "send":  4,
  "check": 5,
  "read":  5,
  "intermediate": 5
}

if (len(sys.argv) != numargs[sys.argv[1]]):
  sys.stderr.write("Usage: NeosClient <send|check|read|ping|intermediate> <resultfilename>  extra parameters\n")
  sys.stderr.write("If 'ping' is selected, no extra parameters are required\n")
  sys.stderr.write("If 'send' is selected, 3rd parameter must be path to the job.xml file\n")
  sys.stderr.write("If 'check', 'read' or 'intermediate' is selected, 3rd and 4th parameter must be the job number and password on NEOS\n")
  sys.exit(1)

neos=xmlrpclib.Server("https://%s:%d" % (NEOS_HOST, NEOS_PORT))
resultfile = open(sys.argv[2], "w")

if sys.argv[1] == "send":
  # Read in XML job file
  xmlfile = open(sys.argv[3], "r")
  xml=""
  buffer=1
  while buffer:
    buffer =  xmlfile.read()
    xml+= buffer
  xmlfile.close()

  try:
    # Send job to NEOS
    (jobNumber,password) = neos.submitJob(xml)
    sys.stdout.write("jobNumber = %d\npassword = %s\n" % (jobNumber,password))
    resultfile.write("jobNumber = %d\npassword = %s\n" % (jobNumber,password))
    resultfile.close()
  # Catch any error and write to file
  except Exception as e:
    resultfile.write("Error: %s" % e.message)
    resultfile.close()

elif sys.argv[1] == 'check':
  (jobNumber, password) = (int(sys.argv[3]), sys.argv[4])
  try:
    resultfile.write(neos.getJobStatus(jobNumber, password))
    resultfile.close()
  # Catch any error and write to file
  except Exception as e:
    resultfile.write("Error: %s" % e.message)
    resultfile.close()

elif sys.argv[1] == "read":
  (jobNumber, password) = (int(sys.argv[3]), sys.argv[4])
  try:
    # Output results
    msg = neos.getFinalResults(jobNumber, password).data

	# If the input was a NL file, write .sol file instead
    if 'Executing NL' in msg:
      results = neos.getOutputFile(jobNumber, password, "ampl.sol").data
      tempFolder = os.path.dirname(sys.argv[2])
      with open(os.path.join(tempFolder, "model.sol"), "w") as fout:
        fout.write(results.replace('\n', '\r\n'))

    else:
      resultfile.write(msg)

    resultfile.close()

  # Catch any error and write to file
  except Exception as e:
    resultfile.write("Error: %s" % e.message)
    resultfile.close()

elif sys.argv[1] == "ping":
  try:
    resultfile.write(neos.ping())
    resultfile.close()
  # Catch any error and write to file
  except Exception as e:
    resultfile.write("Error: %s" % e.message)
    resultfile.close()

elif sys.argv[1] == "intermediate":
  (jobNumber, password) = (int(sys.argv[3]), sys.argv[4])
  try:
    # Output results
    msg = neos.getIntermediateResults(jobNumber, password, 0)[0].data
    resultfile.write(msg)
    resultfile.close()
  # Catch any error and write to file
  except Exception as e:
    resultfile.write("Error: %s" % e.message)
    resultfile.close()
