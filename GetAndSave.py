import os
import openpyxl as pyxl
import shutil
from datetime import datetime
import time
import glob
import zipfile


def move_upload(source_folder, partial='', extention='.*', target_folder='.'):
    '''
    move_upload(source_folder, partial, extention)
    
    This function iterates over all subfolders of <source_folder> and looks for
    files that match <partial>+*+<extention> and move these to <targer_folder>,
    when the extention is <'.zip'> this will be unpacked in <source_folder>
    
    example when uploads will be placed in
    z:\inbox\transfer-yyyy-mm-dd\test-yyyy-mm-dd.zip
    -----------------------------------------------
    source_folder: 'z:/inbox'
    partial: 'test'
    extention: '.zip'
    target_folder: './result'
    
    The above will look for all uploaded files starting matching <test*.zip>,
    copies it to a subfolder, and deletes the source file.
    
    All uploads not matching the pattern will be ignored.
    '''
    if not os.path.exists(source_folder):
        return f'*{source_folder:}* does not exist'
    try:
        moved = []
        if target_folder[-1] == '/':
            target_folder = target_folder[:-1]
        if source_folder[-1] != '/':
            source_folder += '/'
        if not os.path.exists(target_folder):
            os.mkdir(target_folder)
        for item in os.listdir(source_folder):
            if os.path.isdir(source_folder+item):
                for file in glob.glob(source_folder + item + '/'+partial+'*'+extention):
                    shutil.move(file, target_folder+'/')
                    moved.append(file)
                try:
                    os.rmdir(source_folder + '/' + item)
                except:
                    pass
        return moved
    except:
        return 'an error has occurred'
    
    
def find_unpack_zips(target_folder, correction=0):
    '''
    find_unpack_zips takes the arguments:
    - folder: for where the zips
    - correction: to trunc the zip name with more characters than .zip (4) this is optional
    
    If the target directory already exists, it will be deleted.
    
    The zip will be unzipped in a subfolder called <zipfile.zip> minus (4 + correction)
    e.g. find_unpack_zips('./myfolder/', 2)  unzips myzip01.zip in subfolder myzip
    
    When finished unzipping the zipfile is deleted
    
    The function returns a list with processed zipfiles
    
    '''
    files = []
    if target_folder[-1] != '/':
        target_folder += '/'
    path = target_folder
    correction=-4-correction
    for file in glob.glob(path +'/*.zip'):
        source = file
        target = file[:correction]
        if os.path.exists(target):
            shutil.rmtree(target)
                          
        with zipfile.ZipFile(source, 'r') as zip_ref:
            os.mkdir(target)
            zip_ref.extractall(target)
        os.remove(source)
        files.append(file)
    return files


def title_wrapper(func):
    '''
    This function wraps the title with * and -.
    
    Used by print_title
    '''
    def wrapper(*args, **kwargs):
        print('*'*90)
        func(*args, **kwargs)
        print('-'*90)
    return wrapper

@title_wrapper
def print_title(text):
    '''
    This function wraps the text in the title_wrapper
    '''
    print(text)


def choose_retry(text):
    '''
    This generic function returns the input from the user.
    '''
    return input('Please choose Q(uit) or between '+ text + ': ')


def get_all_items(folder, type='folders'):
    '''
    This function gets by default all the subfolders from folder
    If type == files, then it gets all the files in that folder
    '''
    all_outputs = []
    for item in os.listdir(folder):
        if type == 'folders':
            if os.path.isdir(folder+'/' + item): all_outputs.append(item)
        if type == 'files':
            if os.path.isfile(folder+'/' + item): all_outputs.append(item)
    return sorted(all_outputs, reverse=True)


def choose_dir_item(folder, type='folders', what='All'):
    '''
    Pick a subfolder or file from a folder.
    
    '''
    all_items = get_all_items(folder, type)
    if type == 'folders':
        print_title('Choose folder by number')
    else:
        print_title('Choose file by number')
    # create dictionary with all folders
    items = {}
    if what != 'All':
        temp_all_items = []
        for item in all_items:
            if what in item:
                temp_all_items.append(item)
        all_items = temp_all_items

    for count, item in enumerate(all_items):
        items[count+1] = item
        extra =''
        if count+1 < 10: extra = ' '
        print(f'[{count+1}]{extra} {items[count+1]}', end = '\t')
        if (count+1) %3 == 0: print('')
    print('\n')
    item_choosen = False
    while item_choosen == False:
        this_answer = choose_retry(str(1) + ' and ' + str(count + 1))
        try:
            if int(this_answer) > 0 and int(this_answer) <= count + 1:
                item_choosen = True
                return all_items[int(this_answer)-1]
        except:
            if this_answer.lower()=='q':
                print_title('Quiting the notebook run on instruction of the user')
                item_choosen = True
                return ''
            
            
# Functions to evaluate if folders/file exist
def folderfile_exists(folderfile):
    '''
    Pass a single file, folder or list to check if it exists
    Per file, folder it will return True/False and file/folder
    '''
    if not isinstance(folderfile, list):
        folderfile = [folderfile]
    for item in folderfile:
        print(evaluate_folderfile(item))


def evaluate_folderfile(item):
    return str(os.path.exists(item)) + '\t' + item


def output_df(df, target_folder, name, output_type='csv', timestamp=True):
    '''
    Quickly create a CSV, comma separated, of Excel
    
    output_df(df, name, timeestamp)
    df: dataframe to be saved
    name: name of the file
    output_type: xlsx or csv
    timestamp: True or False, default True
    '''
    if not os.path.exists(target_folder):
        os.mkdir(target_folder)
    stamp = time.strftime('%Y%m%d-%H%M%S', time.localtime()) + '-'
    if not timestamp:
        stamp = ''
    if target_folder[-1] != '/':
        target_folder += '/'
    file = target_folder + stamp + name + '.' + output_type
    if output_type == 'csv':
        df.to_csv(file, header=True, index=False, sep=',')
    else:
        df.to_excel(file, index=False)
    print('saved: ' + file)
    return 'saved: ' + file
            

def finished(file):
    '''
    Just outputing that everything is finished.
    '''
    print('=========================================================================')
    print(file + ' has been processed')
    print('=========================================================================')
    choice = input('ENTER to finish').lower()
    print('\n\n')
    print('=========================================================================')
    print('                              Have a nice day')
    print('=========================================================================')