#description: 
# this script updates a path in an excel document 
import openpyxl 
import os
import sys
import ctypes
from ctypes import wintypes


def get_network_path(local_name):
    """
    ------------------------------------------------------------------------------------------------------------
        FUNCTION: Take in a drive letter (ie. 'W:') and return the full network path for that letter.
            Mapped drives are not recognized when trying to open a file, thereby requiring the use of this function
        Parameters:
            local_name str: letter and colon for the mapped drive
        Return str: unc value for the mapped drive
    ------------------------------------------------------------------------------------------------------------
    """
    length = (wintypes.DWORD * 1)()
    result = mpr.WNetGetConnectionW(local_name, None, length)
    if result != ERROR_MORE_DATA:
        raise ctypes.WinError(result)
    remote_name = (wintypes.WCHAR * length[0])()
    result = mpr.WNetGetConnectionW(local_name, remote_name, length)
    if result != ERROR_SUCCESS:
        raise ctypes.WinError(result)
    return remote_name.value


def get_full_path(str_file):
    """
    ------------------------------------------------------------------------------------------------------------
        FUNCTION: determine if a mapped drive file path is being used.
            It then calls a function to replace the drive letter with the full UNC path
        Parameters:
            str_file str: file path that needs to be checked for mapped drives
        Return str: correct file path
    ------------------------------------------------------------------------------------------------------------
    """
    # Check if the file is already a UNC path, if so, return
    if str_file.startswith('\\'):
        return str_file
    # Check to see if the path is a valid file.  If not it is most likely on a mapped drive
    str_file = str_file.replace("'", "")
    file_path = os.path.join(get_network_path(str_file[:2]), str_file[2:])

    return file_path



# Set up variables for getting UNC paths
mpr = ctypes.WinDLL('mpr')
ERROR_SUCCESS = 0x0000
ERROR_MORE_DATA = 0x00EA
wintypes.LPDWORD = ctypes.POINTER(wintypes.DWORD)
mpr.WNetGetConnectionW.restype = wintypes.DWORD
mpr.WNetGetConnectionW.argtypes = (wintypes.LPCWSTR,
                                   wintypes.LPWSTR,
                                   wintypes.LPDWORD)

# Get semi colon separated list of excel files from parameter
xl_param = sys.argv[1]
lst_xl = xl_param.split(';')

# Loop through all Excel files in list
for xl in lst_xl:

    # Get the full unc path if its a named drive
    xl = get_full_path(str_file=xl)
    xl_head = os.path.dirname(xl)

    # Open the workbook
    print(f'Opening workbook: {xl}')
    wb = openpyxl.load_workbook(filename=xl)
    
    # Loop through the sheets within the workbook
    for sheet in wb.worksheets:

        # iterate through all cells containing hyperlinks
        for row in sheet.iter_rows():
            for cell in row:
                # Check if the cell is hyperlinked
                if cell.hyperlink:
                    # Remove the built-in file prefix
                    target = get_full_path(cell.hyperlink.target.replace('file:///',''))

                    # If the hyperlink is to a website, ignore and continue to the next item
                    if target.startswith('https://'):
                        continue

                    # Extract the parts of the hyperlink 
                    target_head = os.path.dirname(target)
                    target_file = os.path.basename(target)

                    # Find the common directories and return the relative path
                    try:
                        rel_head = os.path.relpath(target_head, xl_head)
                    except:
                        rel_head = target_head

                    # Set the new target hyperlink ot the cell
                    new_target = os.path.join(rel_head, target_file)
                    cell.hyperlink.target = new_target
                    print(f'Replaced {target} with {new_target}')

    print('Saving workbook')
    wb.save(xl)


