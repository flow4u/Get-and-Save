'''
Get and Save

This Python script demonstrates the GetAndSave.py capabilities.

Run the Python
'''

import GetAndSave as gs
import pandasas pd
import shutil


# Step 0
'''
Reset the situation
'''

mydirs = ['source', 'test_source - Copy']

for mydir in mydirs:
    try:
        shutil.rmtree(mydir)
    except OSError as e:
        print("Error: %s - %s." % (e.filename, e.strerror))
        
shutil.copytree("./test_source", "./test_source - Copy")


# Step 1
"""
Getting the data from the z:\Inbox after upload to Workspace
Save and extract a zip in the target folder
"""

source_folder = "./test_source - Copy"  # for myDRE replace with: 'z:/inbox'
target_folder = "./source"
partial = "test"
extention = "*.zip"

gs.print_title(
    f"Moving all the uploads from <{source_folder}> to <{target_folder}> matching <{partial}{extention}>"
)
print(gs.move_upload(source_folder, partial, extention, target_folder))



# Step 2
"""
Unpakking all zips
"""

gs.print_title("Unzipping all the files")
gs.find_unpack_zips(target_folder)



# Step 3
"""
Select the correct file and create Panda
"""

source_folder = gs.choose_dir_item(target_folder, "folders", partial)
print('\n'*2)
file = gs.choose_dir_item(target_folder + '/' + source_folder, 'files', partial)

df = pd.read_csv(target_folder + '/' + source_folder + '/'+ file)


# Step Z
"""
Save the Panda
"""

target_folder = "./result"
file_name = "output_file"

gs.print_title("Save as CSV")
gs.output_df(df, target_folder, file_name)
print('\n'*2)
gs.print_title("Save as Excel")
gs.output_df(df, target_folder, file_name, "xlsx")
print('\n'*2)
gs.finished(target_folder + '/' + source_folder + '/'+ file)