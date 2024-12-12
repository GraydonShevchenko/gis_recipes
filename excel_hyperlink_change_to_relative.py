#description: 
# this script updates a path in an excel document 
import openpyxl 
import os
import sys
import ctypes
import arcpy
import logging
from datetime import datetime as dt
from ctypes import wintypes
from argparse import ArgumentParser

def run_app() -> None:
    fld, logger = get_input_parameters()
    replace = ReplaceHyperlinks(folders=fld, logger=logger)
    replace.run_replacements()
    del replace

def get_input_parameters() -> tuple:
    try:
        parser = ArgumentParser(description='This script is used to replace hyperlinks within excel documents '\
                                'with their relative paths.  Specifically meant to be used with AST ')
        parser.add_argument('fld', type=str, help='Semi colon separated list of folders to run')
        parser.add_argument('--log_level', default='INFO', choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'], 
                            help='Log level for message output')
        parser.add_argument('--log_dir', help='Path to the log file directory')

        args = parser.parse_args()
        logger = Environment.setup_logger(args)

        return args.fld, logger

    except Exception as e:
        logging.error(f'Unexpected exception, program terminating: {e.message}')
        raise Exception('Errors exist')




class ArcPyLogHandler(logging.StreamHandler):
    """
    ------------------------------------------------------------------------------------------------------------
        CLASS: Handler used to send logging message to the ArcGIS message window if using ArcMap
    ------------------------------------------------------------------------------------------------------------
    """

    def emit(self, record):
        try:
            msg = record.msg.format(record.args)
        except:
            msg = record.msg

        timestamp = dt.now().strftime('%Y-%m-%d %H:%M:%S')
        if record.levelno == logging.ERROR:
            arcpy.AddError('{} - {}'.format(timestamp, msg))
        elif record.levelno == logging.WARNING:
            arcpy.AddWarning('{} - {}'.format(timestamp, msg))
        elif record.levelno == logging.INFO:
            arcpy.AddMessage('{} - {}'.format(timestamp, msg))

        super(ArcPyLogHandler, self).emit(record)


class Environment:
    """
    ------------------------------------------------------------------------------------------------------------
        CLASS: Contains general environment functions and processes that can be used in python scripts
    ------------------------------------------------------------------------------------------------------------
    """

    def __init__(self):
        pass

    # Set up variables for getting UNC paths
    mpr = ctypes.WinDLL('mpr')

    ERROR_SUCCESS = 0x0000
    ERROR_MORE_DATA = 0x00EA

    wintypes.LPDWORD = ctypes.POINTER(wintypes.DWORD)
    mpr.WNetGetConnectionW.restype = wintypes.DWORD
    mpr.WNetGetConnectionW.argtypes = (wintypes.LPCWSTR,
                                       wintypes.LPWSTR,
                                       wintypes.LPDWORD)

    @staticmethod
    def setup_logger(args):
        """
        ------------------------------------------------------------------------------------------------------------
            FUNCTION: Set up the logging object for message output

            Parameters:
                args: system arguments

            Return: logger object
        ------------------------------------------------------------------------------------------------------------
        """
        log_name = 'main_logger'
        logger = logging.getLogger(log_name)
        logger.handlers = []

        log_fmt = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        log_file_base_name = os.path.basename(sys.argv[0])
        log_file_extension = 'log'
        timestamp = dt.now().strftime('%Y-%m-%d_%H-%M-%S')
        log_file = '{}_{}.{}'.format(timestamp, log_file_base_name, log_file_extension)

        logger.setLevel(args.log_level)

        sh = logging.StreamHandler()
        sh.setLevel(args.log_level)
        sh.setFormatter(log_fmt)
        logger.addHandler(sh)

        if args.log_dir:
            try:
                os.makedirs(args.log_dir)
            except OSError:
                pass
            
            try:
                fh = logging.FileHandler(os.path.join(args.log_dir, log_file))
                fh.setLevel(args.log_level)
                fh.setFormatter(log_fmt)
                logger.addHandler(fh)
                logger.info('Setting up log file')
            except Exception as e:
                logger.removeHandler(fh)

        if os.path.basename(sys.executable).lower() == 'python.exe':
            arc_env = False
        else:
            arc_env = True

        if arc_env:
            arc_handler = ArcPyLogHandler()
            arc_handler.setLevel(args.log_level)
            arc_handler.setFormatter(log_fmt)
            logger.addHandler(arc_handler)

        return logger

    @staticmethod
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
        result = Environment.mpr.WNetGetConnectionW(local_name, None, length)
        if result != Environment.ERROR_MORE_DATA:
            raise ctypes.WinError(result)
        remote_name = (wintypes.WCHAR * length[0])()
        result = Environment.mpr.WNetGetConnectionW(local_name, remote_name, length)
        if result != Environment.ERROR_SUCCESS:
            raise ctypes.WinError(result)
        return remote_name.value

    @staticmethod
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

        if str_file.startswith('\\'):
            return str_file

        # Check to see if the path is a valid file.  If not it is most likely on a mapped drive
        str_file = str_file.replace("'", "")
        if not os.path.isfile(str_file):
            file_path = os.path.join(Environment.get_network_path(str_file[:2]), str_file[2:])
        else:
            file_path = str_file

        return file_path

class ReplaceHyperlinks:
    """
    ------------------------------------------------------------------------------------------------------------
        CLASS: Used to contain the methods for replacing hyperlinks
    ------------------------------------------------------------------------------------------------------------
    """
    def __init__(self, folders: str, logger: logging.Logger) -> None:
        """
        ------------------------------------------------------------------------------------------------------------
            FUNCTION: Class method for initializing the object

            Parameters:
                None

            Return: None
        ------------------------------------------------------------------------------------------------------------
        """
        self.folders = folders.split(';')
        self.xl_files = []
        self.logger = logger

        self.logger.info('Gathering excel files')
        # Get semi colon separated list of excel files from parameter
        for fld in self.folders:
            for root, dirs, files in os.walk(fld):
                for f in files:
                    if f.endswith('xls') or f.endswith('xlsx'):
                        self.xl_files.append(os.path.join(root, f))
    
    def run_replacements(self) -> None:
        """
        ------------------------------------------------------------------------------------------------------------
            FUNCTION: Run through the list of excel files and replace al hyperlinks with relative paths

            Parameters:
                None

            Return: None
        ------------------------------------------------------------------------------------------------------------
        """

        # Loop through all Excel files in list
        for xl in self.xl_files:
        
            # Get the full unc path if its a named drive
            xl = Environment.get_full_path(str_file=xl)
            xl_head = os.path.dirname(xl)

            # Open the workbook
            self.logger.info(f'Opening workbook: {xl}')
            wb = openpyxl.load_workbook(filename=xl)

            # Loop through the sheets within the workbook
            for sheet in wb.worksheets:
            
                # iterate through all cells containing hyperlinks
                for row in sheet.iter_rows():
                    for cell in row:
                        # Check if the cell is hyperlinked
                        if cell.hyperlink:
                            # Remove the built-in file prefix
                            target = cell.hyperlink.target.replace('file:///','')

                            # If the hyperlink is to a website, ignore and continue to the next item
                            if target.startswith('https://'):
                                continue
                            
                            target = Environment.get_full_path(str_file=target)

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
                            self.logger.info(f'Replaced {target} with {new_target}')

            self.logger.info('Saving workbook')
            wb.save(xl)
            wb.close()


if __name__ == '__main__':
    run_app()
