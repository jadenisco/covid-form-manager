import logging
import argparse
import os
import re
import subprocess
import calendar
import shutil
from termcolor import colored
from PyPDF2 import PdfFileReader, PdfFileWriter

# The dictionary of Volunteers
# The Volunteer record we create consists of the following
# number: , last name: , first name:, service dates[]:
volunteers = {}

# System Root directory
# s_root = ['Users',
#          'jdenisco',
#         'Developer',
#         'MrJohnsPython',
#         'mgh-util',
#         'testroot']
# os.sep is used if starting from root

# scratchdir = ['.', 'scratchroot']

scratchdir = ['c:', os.sep,
                'Users', 'JD1060',
                'Developer', 'MrJohnsPython',
                'mgh-util', 'scratchroot']

# testroot = ['.', 'testroot']
# testroot = ['z:', os.sep, 'Developer',
#            'MrJohnsPython',
#            'mgh-util',
#            'testroot']

# testroot = [os.sep, 'Cifs2',
#            'voldept$']

testroot = ['.']

createdfileroot = 'COVID 19 Day Pass - CLEARED FOR WORK'
tmpfilename = 'tmp.pdf'

# Volunteer root directories
adult_volunteer_root = ['.Volunteer Files', 'ADULT MEDICAL AND NONMEDICAL']
junior_volunteer_root = ['.Volunteer Files', 'JUNIOR MEDICAL AND NONMEDICAL']
pet_volunteer_root = ['.Volunteer Files', 'Pet Therapy']
adult_root_dir = ''
junior_root_dir = ''
pet_root_dir = ''
volunteer_dir_db = {}

# Input values
month_on_form = '03'
day_on_form = '01'
year_on_form = '2022'
use_previous_date = False


def _exec_shell_command(command):
    logging.debug("_exec_shell_command({})".format(command))

    prc = subprocess.Popen(command, shell=True, stdin=subprocess.PIPE,
                           stdout=subprocess.PIPE,
                           stderr=subprocess.PIPE)
    ret = prc.wait()
    if ret != 0:
        raise Exception(prc.stderr.read().decode())

    out = prc.stdout.read().decode()
    if out != '':
        print(prc.stdout.read().decode())

    return ret


def _get_filewithpath(path, filename):
    logging.debug("_get_filewithpath({}, {})".format(path, filename))

    pathwithfile = ''
    for pth in path:
        pathwithfile = os.path.join(pathwithfile, pth)
    pathwithfile = os.path.join(pathwithfile, filename)

    return pathwithfile


def read_dict():
    logging.debug('read_dict()')


def read_csv():
    logging.debug('read_csv()')

    with open('./testroot/Volunteer List - Tagged vols.csv', 'r') as f:
        csv = f.readlines()

    v_record = None
    for v in csv:
        vl = v.strip().split(',')
        # print('vl: {}'.format(vl))

        # The record read from the .csv file is consists of:
        # Last Name, Preferred First Name, Status, Number, Date Entered, Service From Date, Service To Date
        if vl[0] != '' and 'Last' not in vl[0]:
            # Start a new record, First save what we have
            if v_record is not None:
                volunteers[v_record['number']] = v_record

            # Then create the new one
            v_record = {'number': int(vl[3]), 'last name': vl[0], 'first name': vl[1], 'service dates': [vl[5]]}
            # print('New Record: {}'.format(v_record))
        else:
            if v_record is not None:
                # Handle additional service dates
                # print('Append Service: {}'.format(vl[5]))
                v_record['service dates'].append(vl[5])

    # Save the last record
    if v_record is not None:
        volunteers[v_record['number']] = v_record

    logging.debug("Volunteers:")
    for v in volunteers.items():
        logging.debug('    {}'.format(v[1]))


def create_validate_forms(create_form):
    logging.debug("create_validate_forms({}): ".format(create_form))

    # Get the root directory
    rootdir = ''
    for sr in testroot:
        rootdir = os.path.join(rootdir, sr)

    # Get the volunteer directory root
    for vd in adult_volunteer_root:
        rootdir = os.path.join(rootdir, vd)

    # For each volunteer get the name
    logging.debug('Volunteer root dir: {}'.format(rootdir))
    for v in volunteers.items():
        v_dir = '{} {}, {}'.format(v[1]['number'], v[1]['last name'], v[1]['first name'])

        # Get the service dates
        for sd in v[1]['service dates']:
            sds = sd.split('-')

            # Get the year
            s_year = os.path.join(v_dir, 'Covid Forms {}'.format(sds[2]))

            # Get the month
            s_month = os.path.join(s_year, calendar.month_name[int(sds[0])])
            s_monthwithroot = os.path.join(rootdir, s_month)
            logging.debug("s_monthwithroot: {}".format(s_monthwithroot))

            # Create or validate the covid form
            fname = '{}.{}.{}.txt'.format(sds[0], sds[1], sds[2])
            fname_with_dir = os.path.join(s_month, fname)
            fname_with_rootdir = os.path.join(rootdir, fname_with_dir)

            logging.debug('file: {}'.format(os.path.join(s_month, fname)))

            if create_form:
                if not os.path.exists(s_monthwithroot):
                    os.makedirs(s_monthwithroot)
                if not os.path.isfile(fname_with_rootdir):
                    with open(fname_with_rootdir, 'w') as f:
                        f.write('I am cleared to work today!\n')
            else:
                if os.path.isfile(fname_with_rootdir):
                    print('The covid form EXISTS for {}'.format(fname_with_dir))
                else:
                    print(colored('The covid form DOES NOT EXIST for {}'.format(fname_with_dir), 'red'))


def _show_pdf(pdffilename):
    logging.debug('_show_pdf({}):'.format(pdffilename))

    if os.name == 'nt':
        _exec_shell_command('start {}'.format(pdffilename))
    elif os.name == 'posix':
        _exec_shell_command('open {}'.format(pdffilename))
    else:
        raise Exception('Unsupported Operating System')


def _create_directory_db(root_dir):
    global volunteer_dir_db

    logging.debug("_build_directory_db({})".format(root_dir))

    for name in os.listdir(root_dir):
        logging.debug("name: {}".format(name))
        namewithpath = os.path.join(root_dir, name)
        logging.debug("namewithpath: {}".format(namewithpath))
        if os.path.isdir(namewithpath):
            vol_num = name.split(' ')[0]
            logging.debug("dir: {} {}".format(vol_num, namewithpath))
            if re.search(r'\d+', vol_num):
                volunteer_dir_db[vol_num] = namewithpath


def _create_volunteer_directory_db():
    global pet_root_dir
    global adult_root_dir
    global junior_root_dir
    global volunteer_dir_db

    logging.debug("_build_volunteer_directory_db()")

    volunteer_dir_db = {}
    _create_directory_db(adult_root_dir)
    _create_directory_db(junior_root_dir)
    _create_directory_db(pet_root_dir)


def _create_volunteer_directory(volunteer_number):
    global volunteer_dir_db
    global adult_root_dir

    logging.debug("_create_volunteer_directory({})".format(volunteer_number))

    firstname = input("Enter the first name on the form: ")
    lastname = input("Enter the last name on the form: ")

    if volunteer_number not in volunteer_dir_db.keys():
        dirname = '{} {}, {}'.format(volunteer_number, lastname, firstname)
        dirname = os.path.join(adult_root_dir, dirname)
        logging.debug("Creating directory: {}".format(dirname))
        if not os.path.exists(dirname):
            os.makedirs(dirname)
            volunteer_dir_db[volunteer_number] = dirname
    else:
        logging.debug('A Directory for {} already exists'.format(volunteer_number))


def _get_pagefilename(pagenumber):
    global month_on_form
    global day_on_form
    global year_on_form
    global createdfileroot
    global use_previous_date

    logging.debug("_get_get_pagefilename()")

    # Get the date to be used in the file name
    if use_previous_date == False:
        answer = input("Enter the Month on the form [{}]: ".format(month_on_form))
        if len(answer) != 0:
            month_on_form = answer

        answer = input("Enter the Day on the form [{}]: ".format(day_on_form))
        if len(answer) != 0:
            day_on_form = answer

        answer = input("Enter the Year on the form [{}]: ".format(year_on_form))
        if len(answer) != 0:
            year_on_form = answer

        answer = input("Do you want to use the same date for every page [n]: ")
        if answer.lower() == 'y':
            use_previous_date = True

    # Get the volunteer number
    answer = input("Enter the volunteer number on the form: ")
    if len(answer) != 0:
        volunteernumber = answer
    else:
        volunteernumber = '({})'.format(pagenumber)

    pagefile = '{} {}_{}_{}-{}.pdf'.format(createdfileroot, month_on_form, day_on_form, year_on_form, volunteernumber)
    pagefilewithpath = _get_filewithpath(scratchdir, pagefile)

    return pagefilewithpath, volunteernumber


def _split_pdf(file_to_split, createdir):
    logging.debug('split_pdf({}):'.format(file_to_split))

    if createdir:
        _create_volunteer_directory_db()

    # Split the file
    use_previous_date = False
    pdf = PdfFileReader(file_to_split)
    for page in range(pdf.getNumPages()):
        pdf_writer = PdfFileWriter()
        pdf_writer.addPage(pdf.getPage(page))

        tmpfile = tmpfilename
        tmpfile = _get_filewithpath(scratchdir, tmpfile)

        logging.debug("Creating a temporary file: {}".format(tmpfile))
        with open(tmpfile, 'wb') as out:
            pdf_writer.write(out)

        _show_pdf(tmpfile)

        pagefilewithpath, volunteer_number = _get_pagefilename(page)

        if createdir:
            _create_volunteer_directory(volunteer_number)
        else:
            print("Creating file: {}, {}".format(volunteer_number, pagefilewithpath))
            if not os.path.isfile(pagefilewithpath):
                os.rename(tmpfile, pagefilewithpath)
            else:
                print("The file {} already exists".format(pagefilewithpath))


def validate_forms(args):
    logging.debug('validate_forms(): ')
    logging.debug(args)

    # Read the data
    read_csv()

    # Parse the csv data and validate the forms
    create_validate_forms(False)


def create_directories(args):
    logging.debug('create_directories()')
    logging.debug(args)

    # Read the data
    read_csv()

    # Parse the csv data and create the directory structure
    create_validate_forms(True)


def _execute_move(src, dst):
    logging.debug("_execute_move({}, {})".format(src, dst))

    if not os.path.isdir(dst):
        print("ERROR: The Directory {} does NOT exist.")
        return

    cv_date = re.search(r'(''|[0-1])[0-9]_(''|[0-3])[0-9]_20\d{2}', src).group().split('_')
    cv_month = calendar.month_name[int(cv_date[0])]
    cv_year = cv_date[2]

    cv_form_dir = 'Covid Forms {}'.format(cv_year)
    cv_form_dir = os.path.join(dst, cv_form_dir)
    if not os.path.isdir(cv_form_dir):
        os.makedirs(cv_form_dir)

    cv_form_dir = os.path.join(cv_form_dir, cv_month)
    if not os.path.isdir(cv_form_dir):
        os.makedirs(cv_form_dir)

    src_len = len(src.split(os.sep))
    filename = src.split(os.sep)[src_len-1]
    filewithpath = os.path.join(cv_form_dir, filename)
    if not os.path.isfile(filewithpath):
        shutil.move(src, cv_form_dir)
    else:
        print("\nThe file {} already exists".format(filewithpath))


def move(args):
    """
    Move the files that were split using the split command to the appropriate directories.
    
    :param args: The parsed input arguments
    :type args: Namespace

    Example: python mgh-util.py move

    """
    global scratchdir
    global volunteer_dir_db

    logging.debug("move({})".format(args))

    _create_volunteer_directory_db()

    scratchpath = ''
    for sp in scratchdir:
        scratchpath = os.path.join(scratchpath, sp)

    for name in os.listdir(scratchpath):
        src = os.path.join(scratchpath, name)
        if not os.path.isfile(src):
            logging.debug("{} is not a file.".format(src))
            continue

        logging.debug("src: {}".format(src))

        if re.search(r'CLEARED FOR WORK (''|[0-1])[1-9]_(''|[0-3])[0-9]_20\d{2}-\d+', src):
            vol_num = re.search(r'-\d+.pdf', src).group().lstrip('-').rstrip('.pdf')
            logging.debug("Vol Number: {}".format(vol_num))

            if vol_num in volunteer_dir_db.keys():
                dst = volunteer_dir_db[vol_num]
                print("Moving file from: {}".format(src))
                print("To: {}".format(dst))
                # ans = input("Is this ok y/n [n]?  ")
                # ans.lower()
                # if ans == 'y':
                _execute_move(src, dst)
            else:
                print("The Directory for {} was not found.".format(src))
        else:
            print("The file {} was not moved.".format(src))


def split(args):
    """
    Split a single pdf that contains covid attestation forms. The single page file names will
    contain the date and volunteer number. If the page can not be read, the page number is in
    place of the volunteer number.

    :param args: The parsed input arguments
    :type args: Namespace

    Example: python mgh-util.py split ./scratchroot/covid-forms-Dec-14-15.pdf

    When the command is run:

    1. The first page contained within the pdf is shown.
    2. The user then enters the date and volunteer number.
    3. The file is then renamed to reflect the entered information
    4. The previous steps are repeated for all the pages contained within the original pdf.

    """

    logging.debug("split_move({})".format(args))

    filetobesplit = args.filename
    logging.debug("filetobesplit: {}".format(filetobesplit))

    # Get some input
    _split_pdf(filetobesplit, args.create)


def mgh_util(args):
    global adult_root_dir
    global junior_root_dir
    global pet_root_dir

    if args.debug:
        logging.basicConfig(level=logging.DEBUG)

    logging.debug("mgh_util({})".format(args))

    # Get the volunteer root directories
    t_root = ''
    for tr in testroot:
        t_root = os.path.join(t_root, tr)

    adult_root_dir = t_root
    for avr in adult_volunteer_root:
        adult_root_dir = os.path.join(adult_root_dir, avr)

    junior_root_dir = t_root
    for jvr in junior_volunteer_root:
        junior_root_dir = os.path.join(junior_root_dir, jvr)

    pet_root_dir = t_root
    for pvr in pet_volunteer_root:
        pet_root_dir = os.path.join(pet_root_dir, pvr)

    if args.func:
        return args.func(args)
    else:
        return 0


if __name__ == '__main__':
    main_parser = argparse.ArgumentParser(
        prog='mgh-util',
        description='These utilities can be used by the MGH volunteer department.',
        epilog='See "%(prog)s help COMMAND" for help on a specific command.')
    main_parser.add_argument('--debug', '-d', action='count', help='Print debug output')
    main_parser.add_argument('--create', '-c', action='count', help='Create a directory if needed')
    sub_parsers = main_parser.add_subparsers()

    cr_parser = sub_parsers.add_parser('create-dirs', help='Create the Volunteer Directory Structure')
    cr_parser.set_defaults(func=create_directories)

    v_parser = sub_parsers.add_parser('validate', help='Validate the Volunteer Directory Structure')
    v_parser.set_defaults(func=validate_forms)

    sm_parser = sub_parsers.add_parser('split', help='Split the scanned pdf file')
    sm_parser.add_argument('filename', help='The pdf file to be split and moved')
    sm_parser.set_defaults(func=split)

    mv_parser = sub_parsers.add_parser('move', help='Move the files')
    mv_parser.set_defaults(func=move)

    main_args = main_parser.parse_args()
    mgh_util(main_args)
