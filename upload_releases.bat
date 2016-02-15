:: Uploads all zip files in the Releases directory to Sourceforge
:: Requires 'pscp.exe' on the system path, installed with PuTTY

set /p user="Enter sourceforge username (e.g. jackdunnnz): "

set start=pscp
set end=%user%@frs.sourceforge.net:/home/frs/project/opensolver

:: Upload each zip file in releases
for /R %%G in (*.zip) do %start% "%%G" %end%

echo Upload complete! Remember to enable 'exclude stats' for readme.txt in the web interface. 