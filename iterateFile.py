# import os
# import pandas as pd
# from datetime import datetime

# # Directory path
# dir_path = 'Z:\DataBC_Contractors\SRE'

# # List to store file data
# file_data = []

# # Iterate over all the files in directory
# for file_name in os.listdir(dir_path):
#     file_info = os.stat(os.path.join(dir_path, file_name))
    
#     # Append file data to list
#     file_data.append({
#         'file_name': file_name,
#         'last_modified': datetime.fromtimestamp(file_info.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
#     })


# with open("../file_info1.csv", 'w',newline='') as f:
#     writer = csv.writer(f)
#     header=['Category','Location','Description','Link','Last','Updated Date','Health','Status','Owner','Notes']
#     allGroup=[]
#     with open("GroupName.txt", 'r') as t:
#         for line in t:
#             # print(line)
#             header.append(line.strip())
#             allGroup.append(line.strip())
#     # print(header)
#     writer.writerow(header)


# # Create a DataFrame from the file data
# df = pd.DataFrame(file_data)
# print(df)
# # Write DataFrame to Excel
# df.to_excel('./file_info.xlsx', index=False)

# import os
# import pandas as pd
# from datetime import datetime

# # Directory path
# dir_path = 'Z:\DataBC_Contractors\SRE'
# count=0
# # List to store file data
# file_data = []
# # Iterate over all the files in directory and subdirectories
# for root, dirs, files in os.walk(dir_path):
#     for file_name in files:
#         count+=1
#         print('No-'+str(count)+''+file_name)
#         file_info = os.stat(os.path.join(root, file_name))
#         file_path = os.path.join(root, file_name)
#         current_dir_name = os.path.basename(root)
#         file_name_without_ext = os.path.splitext(file_name)[0]
#         # Append file data to list
#         file_data.append({
#             'Category': ' ',  # Update as needed
#             'Location': current_dir_name,
#             'Description': file_name_without_ext,  # Update as needed
#             'Link': '\\\\sfp.idir.bcgov\\s177\\S7793\\' + file_path[3:],
#             'Last Updated Date': datetime.fromtimestamp(file_info.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
#             'Health': ' ',  # Update as needed
#             'Status': 'Valid',  # Update as needed
#             'Owner': ' ',  # Update as needed
#             'Notes': ' '  # Update as needed
#         })

# # Create a DataFrame from the file data
# df = pd.DataFrame(file_data)

# # Write DataFrame to Excel
# df.to_excel('./Iterate/file_info1.xlsx', index=False)

import os
import pandas as pd
from datetime import datetime
import argparse
import win32wnet

# Set up command line argument parsing
parser = argparse.ArgumentParser(description='A script for processing directory paths.')
parser.add_argument('-p', '--path', type=str, help='The input directory path. Default is "Z:\\DataBC_Contractors\\SRE".')

args = parser.parse_args()

# Directory path
dir_path = args.path if args.path else 'Z:\\DataBC_Contractors\\SRE'
count=0
# List to store file data
file_data = []

# Iterate over all the files in directory and subdirectories
for root, dirs, files in os.walk(dir_path):
    for file_name in files:
        count+=1
        print('No-'+str(count)+''+file_name)
        file_info = os.stat(os.path.join(root, file_name))
        file_path = os.path.join(root, file_name)
        current_dir_name = os.path.basename(root)
        file_name_without_ext = os.path.splitext(file_name)[0]
        
        # Get the network share path
        try:
            network_share_path = win32wnet.WNetGetUniversalName(root, 1)
            file_link = network_share_path+'\\'+file_name
        except Exception as e:
            print(f"Could not get network share path for {root}. Error: {e}")
            file_link = file_path  # fallback to local path if network share path could not be obtained
        
        # Append file data to list
        file_data.append({
            'Category': ' ',  # Update as needed
            'Location': current_dir_name,
            'Description': file_name_without_ext,  # Update as needed
            'Link': file_link,
            'Last Updated Date': datetime.fromtimestamp(file_info.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
            'Health': ' ',  # Update as needed
            'Status': 'Valid',  # Update as needed
            'Owner': ' ',  # Update as needed
            'Notes': ' '  # Update as needed
        })

# Create a DataFrame from the file data
df = pd.DataFrame(file_data)

# Write DataFrame to Excel
df.to_excel('file_info1.xlsx', index=False)
