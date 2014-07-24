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

if len(sys.argv) != 3:
  sys.stderr.write("Usage: NeosClient <xmlfilename> <resultfilename>\n")
  sys.exit(1)

neos=xmlrpclib.Server("http://%s:%d" % (NEOS_HOST, NEOS_PORT))

# Read in XML job file
xmlfile = open(sys.argv[1], "r")
xml=""
buffer=1
while buffer:
  buffer =  xmlfile.read()
  xml+= buffer
xmlfile.close()

resultfile = open(sys.argv[2], "w")
try:
  # Send job to NEOS
  (jobNumber,password) = neos.submitJob(xml)
  sys.stdout.write("jobNumber = %d\tpassword = %s\n" % (jobNumber,password))

  # Wait for completion
  offset=0
  status="Waiting"
  while status == "Running" or status=="Waiting":
    time.sleep(1)
    msg, offset = neos.getIntermediateResults(jobNumber,password,offset)
    sys.stdout.write(msg.data)
    status = neos.getJobStatus(jobNumber, password)

  # Output results
  msg = neos.getFinalResults(jobNumber, password).data
  resultfile.write(msg)
  resultfile.close()

# Catch any error and write to file
except Exception as e:
  resultfile.write("Error: %s" % e.message)
  resultfile.close()
