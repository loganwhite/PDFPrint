import os
import sys
import getopt
import win32api
import win32print


def init_printer():
    name = win32print.GetDefaultPrinter()

    print "printer name: " + name

    # printdefaults = {"DesiredAccess": win32print.PRINTER_ACCESS_ADMINISTER}
    printdefaults = {"DesiredAccess": win32print.PRINTER_ACCESS_USE}
    handle = win32print.OpenPrinter(name, printdefaults)

    level = 2
    attributes = win32print.GetPrinter(handle, level)

    print "Old Duplex = %d" % attributes['pDevMode'].Duplex

    # attributes['pDevMode'].Duplex = 1    # no flip
    attributes['pDevMode'].Duplex = 2    # flip up
    # attributes['pDevMode'].Duplex = 3  # flip over

    ## 'SetPrinter' fails because of 'Access is denied.'
    ## But the attribute 'Duplex' is set correctly
    try:
        win32print.SetPrinter(handle, level, attributes, 0)
    except:
        print "win32print.SetPrinter: set 'Duplex'"

    return name, printdefaults, handle, attributes


def clean_printer(handle):
    win32print.ClosePrinter(handle)


def print_pdf(filename):
    print "print file: " + filename
    res = win32api.ShellExecute(0, 'print', filename, None, '.', 0)


def traversal_dir(dirname):
    if not os.path.exists(dirname):
        print "path does not exist!"
        return
    if not os.path.isdir(dirname):
        print "not a path!"
        return

    file_list = [fname for fname in os.listdir(dirname)]

    for filename in file_list:
        absolute_filename = dirname + os.sep + filename
        if os.path.isdir(absolute_filename):
            traversal_dir(absolute_filename)
        else:
            if os.path.isfile(absolute_filename) and filename[-4:] == '.pdf':
                print_pdf(absolute_filename)


def main():
    name, printdefaults, handle, attributes = init_printer()
    try:
        opts, args = getopt.getopt(sys.argv[1:], "d:", ["d="])
    except getopt.GetoptError as err:
        # print help information and exit:
        print str(err)  # will print something like "option -a not recognized"
        print "d: d="
        sys.exit(2)

    for o, value in opts:
        if o in ('-d', "--dir"):
            directory = value
            traversal_dir(directory)

        else:
            assert False, "unhandled option"

    clean_printer(handle)






if __name__ == '__main__':
    main()
