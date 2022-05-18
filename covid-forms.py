import logging
import argparse
import os
import re
import subprocess
import calendar
import email
from email.header import decode_header
import shutil
from termcolor import colored
from PyPDF2 import PdfFileReader, PdfFileWriter

# The dictionary of Volunteers
# The Volunteer record we create consists of the following
# number: , last name: , first name:, service dates[]:
volunteers = {}

vol_root_dir = '/Users/jdenisco/Developer/Windows/testroot'
script_dir = vol_root_dir + '/scripts/cfm-mac/covid-form-manager'
forms_dir = script_dir + '/forms'

# Volunteer root directories
adult_volunteer_root_dir = vol_root_dir + '/.Volunteer Files/ADULT MEDICAL AND NONMEDICAL'
junior_volunteer_root_dir = vol_root_dir + '/.Volunteer Files/JUNIOR MEDICAL AND NONMEDICAL'
pet_volunteer_root_dir = None
# pet_volunteer_root_dir = vol_root_dir + '/.Volunteer Files/Pet Therapy'

# System Root directory
# s_root = ['Users',
#          'jdenisco',
#         'Developer',
#         'MrJohnsPython',
#         'mgh-util',
#         'testroot']
# os.sep is used if starting from root

#testroot = ''

#scratch_dir = './forms'

#scratchdir = ['c:', os.sep,
#                'Users', 'JD1060',
#                'Developer', 'MrJohnsPython',
#                'mgh-util', 'scratchroot']

#testroot = ['.', 'testroot']
# testroot = ['z:', os.sep, 'Developer',
#            'MrJohnsPython',
#            'mgh-util',
#            'testroot']
#

# testroot = [os.sep, 'Cifs2',
#            'voldept$']

# testroot = ['.']

created_file_root = 'COVID 19 Day Pass - CLEARED FOR WORK'
tmp_filename = 'tmp.pdf'

#adult_root_dir = ''
#junior_root_dir = ''
#pet_root_dir = ''
volunteer_dir_db = {}
volunteer_name_dir_db = {}

# Input values
month_on_form = '12'
day_on_form = '15'
year_on_form = '2021'
use_previous_date = False

def _ask_y_n(question, default='y'):

    while True:
        answer = input("{} [{}]?: ".format(question, default))
        if answer == '':
            answer = default
        answer = answer.lower()
        if answer == 'y' or answer == 'n':
            break
        print("Invalid input, Please enter y or n.")

    return answer

def _exec_shell_command(command):
    logging.debug("_exec_shell_command({})".format(command))

    prc = subprocess.Popen(command, shell=True, stdin=subprocess.PIPE,
                           stdout=subprocess.PIPE,
                           stderr=subprocess.PIPE)

    err = ''
    out = prc.stdout.read().decode()
    if out != '':
        logging.debug("_exec_shell_command: {}".format(out))
    else:
        err = prc.stderr.read().decode()
        if err != '':
            logging.error("_exec_shell_command: {}".format(err))

    return [out, err]

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


def _show_pdf(pdf_filename):
    logging.debug('_show_pdf({}):'.format(pdf_filename))

    if os.name == 'nt':
        _exec_shell_command('start {}'.format(pdf_filename))
    elif os.name == 'posix':
        _exec_shell_command('open {}'.format(pdf_filename))
    else:
        raise Exception('Unsupported Operating System')


def _create_name_directory_db(root_dir):
    global volunteer_name_dir_db

    logging.debug("_create_name_directory_db({})".format(root_dir))

    for name in os.listdir(root_dir):
        namewithpath = os.path.join(root_dir, name)
        if os.path.isdir(namewithpath):
            first_name = name.split(' ')[2]
            last_name = name.split(' ')[1].rstrip(',')
            vol_num = name.split(' ')[0]
            key = ''.join([first_name, last_name]).lower()
            logging.debug("entry: {} {}".format(key, namewithpath))
            if key in volunteer_name_dir_db.keys():
                volunteer_name_dir_db[key].append(namewithpath)
            else:
                db_entry_list = []
                db_entry_list.append(namewithpath)
                volunteer_name_dir_db[key] = db_entry_list


def _create_volunteer_name_directory_db():
    global adult_root_dir
    global junior_root_dir
    global volunteer_name_dir_db

    logging.debug("_create_volunteer_name_directory_db()")

    volunteer_name_dir_db = {}
    _create_name_directory_db(adult_root_dir)
    _create_name_directory_db(junior_root_dir)


def _create_directory_db(root_dir):
    global volunteer_dir_db

    logging.debug("_build_directory_db({})".format(root_dir))

    if not root_dir:
        return

    for name in os.listdir(os.path.abspath(root_dir)):
        namewithpath = os.path.join(root_dir, name)
        logging.debug("namewithpath: {}".format(namewithpath))
        if os.path.isdir(namewithpath):
            vol_num = name.split(' ')[0]
            logging.debug("dir: {} {}".format(vol_num, namewithpath))
            if re.search(r'\d+', vol_num):
                volunteer_dir_db[vol_num] = namewithpath


def _create_volunteer_directory_db():
    global adult_volunteer_root_dir
    global junior_volunteer_root_dir
    global pet_volunteer_root_dir
    global volunteer_dir_db

    logging.debug("_build_volunteer_directory_db()")

    volunteer_dir_db = {}
    _create_directory_db(adult_volunteer_root_dir)
    _create_directory_db(junior_volunteer_root_dir)
    _create_directory_db(pet_volunteer_root_dir)


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


def _get_page_filename(page_number):
    global month_on_form
    global day_on_form
    global year_on_form
    global created_file_root
    global use_previous_date

    logging.debug("_get_get_page_filename()")

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
        volunteernumber = '({})'.format(page_number)

    pagefile = '{} {}_{}_{}-{}.pdf'.format(createdfileroot, month_on_form, day_on_form, year_on_form, volunteernumber)
    pagefilewithpath = _get_filewithpath(scratchdir, pagefile)

    return pagefilewithpath, volunteernumber


def _split_pdf(file_to_split, create_dir):
    logging.debug('split_pdf({}):'.format(file_to_split))

    if create_dir:
        _create_volunteer_directory_db()

    # Split the file
    use_previous_date = False
    pdf = PdfFileReader(file_to_split)
    for page in range(pdf.getNumPages()):
        pdf_writer = PdfFileWriter()
        pdf_writer.addPage(pdf.getPage(page))

        tpf = os.path.join(os.path.abspath(forms_dir), tmp_filename)
        logging.debug("Creating a temporary file: {}".format(tpf))
        with open(tpf, 'wb') as out:
            pdf_writer.write(out)

        _show_pdf(tpf)

        # jadfix: start here
        pagefilewithpath, volunteer_number = _get_page_filename(page)

        if create_dir:
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

    Example: python covid-forms.py move

    """
    global scratch_dir
    global volunteer_dir_db

    logging.debug("move({})".format(args))

    _create_volunteer_directory_db()

    # Move the files
    fd = os.path.abspath(forms_dir)
    for name in os.listdir(fd):
        src = os.path.join(fd, name)
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

def _get_file():
    global forms_dir

    logging.debug("getfile()")

    fd = os.path.abspath(forms_dir)
    for fname in os.listdir(fd):
        file_with_path = os.path.join(fd, fname)
        if os.path.isfile(file_with_path):
            _show_pdf(file_with_path)
            answer = _ask_y_n("Do you want to split the file {}".format(fname), default='n')
            if answer == 'y':
                return file_with_path

    return None

def _rename_file(src):
    logging.debug("_rename_file({})".format(src))

    print("The file we are going to split is: {}".format(src))
    dst = input("What would you like to rename it to: ")
    if os.sep not in dst:
        dst = os.path.join(os.path.dirname(src), dst)
    ds = dst.split('.')
    if len(ds) > 0 and ds[len(ds)-1] != 'pdf':
        dst = dst + '.pdf'
    
    logging.debug("The new filename is: {}".format(dst))
    os.rename(src, dst)
    return dst

def split(args):
    """
    Split a single pdf that contains covid attestation forms. The single page file names will
    contain the date and volunteer number. If the page can not be read, the page number is in
    place of the volunteer number.

    :param args: The parsed input arguments
    :type args: Namespace

    Example: python covid-forms.py split ./scratchroot/covid-forms-Dec-14-15.pdf

    When the command is run:

    1. The first page contained within the pdf is shown.
    2. The user then enters the date and volunteer number.
    3. The file is then renamed to reflect the entered information
    4. The previous steps are repeated for all the pages contained within the original pdf.

    """

    logging.debug("split({})".format(args))

    if args.file:
        file_to_split = args.file
    else:
        file_to_split = _get_file()
    if file_to_split == None:
        logging.error("There is not a valid file to split!")

    logging.debug("file_to_split: {}".format(file_to_split))

    answer = _ask_y_n("Would you like to rename the file: {}".format(os.path.basename(file_to_split)))
    if(answer == 'y'):
        file_to_split = _rename_file(file_to_split)

    # Get the current association with .pdf
    # and change it to MSEdgePDF
    pdf_type = ''
    if os.name == 'nt':
        out, err = _exec_shell_command('assoc .pdf')
        if out != '':
            out = out.rstrip().split('=')
            if len(out) == 2:
                pdf_type = out[1]
                _exec_shell_command('assoc .pdf=MSEdgePDF')

    # Get some input
    _split_pdf(file_to_split, args.create)

    # Change the associate for .pdf back
    if os.name == 'nt':
        _exec_shell_command('assoc .pdf={}'.format(pdf_type))

def _extract_name_date(msg):
    logging.debug("_extract_name_date(msg)")

    content_type = msg.get_content_type()
    logging.debug("Type: {}".format(content_type))

    # Use a plain section
    if content_type == 'text/plain':
        try:
            body = msg.get_payload(decode=True)
            if body:
                db = body.decode()
                # Look through the body for a date, time and name
                pattern = re.findall(r'\d+/\d+/\d+ \d+:\d+:\d+\r\n[a-zA-Z]+ [a-zA-Z]+\r\n', db)
                if pattern:
                    sp = pattern[0].split('\r\n')
                    clear_date = sp[0].split(" ")[0]
                    name = sp[1]
                    logging.debug("The date {} and name {} was found".format(clear_date, name))
                    return clear_date, name
        except:
            logging.error("The body of the message could not be decoded")
            return None, None
    return None, None

def _move_email(src, clear_date, name):
    global createdfileroot
    global volunteer_name_dir_db

    logging.debug("_move_email({}, {}, {})".format(src, clear_date, name))
    key = name.replace(' ', '').lower()
    if key in volunteer_name_dir_db.keys():
        if len(volunteer_name_dir_db[key]) > 1:
            print(colored('There is more than one directory for {}: {}'.format(key, volunteer_name_dir_db[key]), 'red'))
        else:
            # Rename the file
            dst = volunteer_name_dir_db[key][0]
            vol_num = os.path.basename(dst).split(' ')[0]
            newfile = "{} {}-{}.eml".format(createdfileroot, clear_date.replace('/', '_'), vol_num)
            newfilewithpath = os.path.join(os.path.dirname(src), newfile)
            if src != newfilewithpath:
                os.rename(src, newfilewithpath)
            _execute_move(newfilewithpath, dst)

    else:
        print(colored('A directory is not found for {}'.format(key), 'red'))


def _read_emails(filename):
    logging.debug("read_emails({})".format(filename))

    with open(filename, 'rb') as f:
        msg = email.message_from_bytes(f.read())

    for subject, encoding in decode_header(msg['Subject']):
        if isinstance(subject, bytes):
            subject = subject.decode(encoding)
        content_type = msg.get_content_type()
        logging.debug("Subject: {} Encoding: {}".format(subject, encoding))
        logging.debug("Type: {}".format(content_type))

    if msg.is_multipart():
        for part in msg.walk():
            clear_date, name = _extract_name_date(part)
            if clear_date:
                break
    else:
        clear_date, name = _extract_name_date(msg)

    if clear_date:
        _move_email(filename, clear_date, name)
 
    
def read_emails(args):
    logging.debug("email({})".format(args))

    _create_volunteer_name_directory_db()

    scratchpath = ''
    for sp in scratchdir:
        scratchpath = os.path.join(scratchpath, sp)

    for name in os.listdir(scratchpath):
        filewithpath = os.path.join(scratchpath, name)
        if filewithpath.find('.eml') == -1:
            logging.debug("{} is not an email.".format(filewithpath))
            continue

        _read_emails(filewithpath)


def mgh_util(args):
    global adult_root_dir
    global junior_root_dir
    global pet_root_dir

    if args.debug:
        logging.basicConfig(level=logging.DEBUG)

    logging.debug("mgh_util({})".format(args))

    """
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
    """

    if args.func:
        return args.func(args)
    else:
        return 0


if __name__ == '__main__':
    main_parser = argparse.ArgumentParser(
        prog='covid-forms',
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
    sm_parser.add_argument('--file', '-f', help='The pdf file to be split and moved')
    sm_parser.set_defaults(func=split)

    sm_parser = sub_parsers.add_parser('read-emails', help='read emails')
    sm_parser.set_defaults(func=read_emails)

    mv_parser = sub_parsers.add_parser('move', help='Move the files')
    mv_parser.set_defaults(func=move)

    main_args = main_parser.parse_args()
    mgh_util(main_args)
