The following script can download attachment from either eml file in this folder or directly from outlook app and 
will put it in separate folder based on Nota Dinas number and update the summary.
To run this script make sure python3 & necessary library is installed, such as:
	1. email
	2. re
	3. shutil
	4. pandas
	5. pywin32
When requirement is complete, make sure previous excel summary of nodin is in place.
How to run it:
	1. This script will download attahcment from eml file
		python get_attachment_eml.py
	2. This script will download attahcment from outlook application, make sure your outlook is openend and already refreshed
		python get_attachment_direct.py