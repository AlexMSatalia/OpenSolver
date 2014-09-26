#!/usr/bin/env python

# Python script for contacting NEOS over XML-RPC
# NeosClient.py is included in NEOS download package:
# http://www.neos-server.org/neos/downloads.html
# Modifications by JWD

import sys
import xmlrpclib
import time

NEOS_HOST="www.neos-server.org"
NEOS_PORT=3332

if len(sys.argv) < 4 or (sys.argv[1] != "send" and len(sys.argv) != 5):
  sys.stderr.write("Usage: NeosClient <send|check|read> <resultfilename>  extra parameters\n")
  sys.stderr.write("If 'send' is selected, 3rd parameter must be path to the job.xml file\n")
  sys.stderr.write("If 'check' or 'read' is selected, 3rd and 4th parameter must be the job number and password on NEOS\n")
  sys.exit(1)

neos=xmlrpclib.Server("http://%s:%d" % (NEOS_HOST, NEOS_PORT))
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
    resultfile.write(msg)
    resultfile.close()
  # Catch any error and write to file
  except Exception as e:
    resultfile.write("Error: %s" % e.message)
    resultfile.close()
