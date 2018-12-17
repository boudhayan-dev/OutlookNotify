import subprocess
subprocess.check_output("cd path_to_batch_file_escaping_the_back_slash file && drive_letter: && name_of_the_batch_file.bat",stderr=subprocess.STDOUT,shell=True)
subprocess.check_output("schtasks /Create /tn mytask12 /sc ONSTART /st time_in_24_hour_format /tr absolute_path_of_batchfile_with_backslash_escaped.bat",stderr=subprocess.STDOUT)

