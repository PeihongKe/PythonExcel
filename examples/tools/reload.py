"""
PyXLL Examples: reload.py

This script can be called from outside of Excel to load and
reload modules using PyXLL.

It uses win32com (part of pywin32) to call into Excel to two built-in
PyXLL Excel macros ('pyxll_reload' and 'pyxll_rebind') and another
macro 'pyxll_import_file' defined in this file.

The PyXLL reload and rebind commands are only available in developer mode,
so ensure that developer_mode in the pyxll.cfg configuration is set to 1.

Excel must already be running for this script to work.

Example Usage:

# reload all modules
python reload.py

# reload a specific module
python reload.py <filename>
"""
import sys
import os
import pickle
import logging
import imp

_log = logging.getLogger(__name__)

def main():
    # pywin32 must be installed to run this script
    try:
        import win32com.client
    except ImportError:
        _log.error("*** win32com.client could not be imported          ***")
        _log.error("*** tools.reload.py will not work                  ***")
        _log.error("*** to fix this, install the pywin32 extensions.   ***")
        return -1

    # any arguments are assumed to be filenames
    # of modules to reload
    filenames = None
    if len(sys.argv) > 1:
        filenames = sys.argv[1:]

    # this will fail if Excel isn't running
    xl_app = win32com.client.GetActiveObject("Excel.Application")

    # load the modules listed on the command line by
    # calling the macro defined in this file.
    if filenames:
        for filename in filenames:
            filename = os.path.abspath(filename)
            print("re/importing %s" % filename)
            response = xl_app.Run("pyxll_import_file", filename)
            response = pickle.loads(str(response))
            if isinstance(response, Exception):
                raise response

        # once all the files have been imported or reloaded
        # call the built-in pyxll_rebind macro to update the
        # Excel functions without reloading anything else
        xl_app.Run("pyxll_rebind")
        print("Rebound PyXLL functions")

    else:
        # call the built-in pyxll__reload macro
        xl_app.Run("pyxll_reload")
        print("Reloaded all PyXLL modules")

#
# in order to be able to reload particular files we add
# an Excel macro that has to be loaded by PyXLL
#
try:
    from pyxll import xl_macro

    @xl_macro
    def pyxll_import_file(filename):
        """
        imports or reloads a python file.

        Returns an Exception on failure or True on success
        as a pickled string.
        """
        # keep a copy of the path to restore later
        sys_path = list(sys.path)
        try:
            # insert the path to the pythonpath
            path = os.path.dirname(filename)
            sys.path.insert(0, path)

            try:
                # try to load/reload the module
                basename = os.path.basename(filename)
                modulename, ext = os.path.splitext(basename)
                if modulename in sys.modules:
                    module = sys.modules[modulename]
                    imp.reload(module)
                else:
                    __import__(modulename)

            except Exception as e:
                # return the pickled exception
                return pickle.dumps(e)

        finally:
            # restore the original path
            sys.path = sys_path

        return pickle.dumps(True)

except ImportError:
    pass

if __name__ == "__main__":
    sys.exit(main())