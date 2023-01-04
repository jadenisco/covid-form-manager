import logging
import argparse
import os
import re
import subprocess
import calendar
import json
import time
import glob
from msg_parser import MsOxMessage
import email
from email.header import decode_header
import shutil
from PyPDF2 import PdfFileReader, PdfFileWriter

#if os.name == 'nt': 
#    import pywintypes
#    import win32gui
#    WAIT_FOR_SHOW_SECS = 5

# This is a change
# The dictionary of Volunteers
# The Volunteer record we create consists of the following
# number: , last name: , first name:, service dates[]:
volunteers = {}

vol_root_dir = '//Cifs2/voldept$'
# vol_root_dir = '/Users/jdenisco/Developer/Windows/testroot'
# vol_root_dir = 'z:/Developer/Windows/testroot'
script_dir = vol_root_dir + '/scripts/cfm-test/covid-form-manager'
# script_dir = vol_root_dir + '/scripts/cfm-mac/covid-form-manager'
forms_dir = script_dir + '/forms'

# jadfix: Should use only the forms directory, don't need this
emails_dir = './emails'
# emails_dir = './forms-01'

# Volunteer root directories
adult_volunteer_root_dir = vol_root_dir + '/.Volunteer Files/ADULT MEDICAL AND NONMEDICAL'
junior_volunteer_root_dir = vol_root_dir + '/.Volunteer Files/JUNIOR MEDICAL AND NONMEDICAL/Active JR Volunteers'
# pet_volunteer_root_dir = None
pet_volunteer_root_dir = vol_root_dir + '/.Volunteer Files/Pet Therapy'

dirs_to_search = [adult_volunteer_root_dir, junior_volunteer_root_dir, pet_volunteer_root_dir]
name_db_filename = script_dir + '/name_db.json'
num_db_filename = script_dir + '/num_db.json'
patch_db_filename = script_dir + '/patch_db.json'

dry_run = False
tmp_filename = 'tmp.pdf'

volunteer_name_db = {}
volunteer_num_db = {}
patch_db = {}

# jadfix: Don't need these
dup_volunteers_db = {}
volunteer_dir_db = {}

month_on_form = '01'
day_on_form = '02'
year_on_form = '2023'
use_previous_date = False

def _ask_y_n(question, default='y'):

    while True:
        answer = input("{} [{}]? ".format(question, default))
        if answer == '':
            answer = default
        answer = answer.lower()
        if answer == 'y' or answer == 'n':
            break
        print("Invalid input, Please enter y or n.")

    return answer

def _ask_value(question, min, max, default_value):

    while True:
        if default_value == None:
            answer = input("{}? ".format(question))
            if answer == '':
                return ''
        else:
            answer = input("{} [{}]? ".format(question, default_value))
            if answer == '':
                answer = default_value

        try:
            v = int(answer)
        except:
            print("Please enter a valid number.")
            continue

        if v in range(min, max+1):
            if v in  range(0, 9+1):
                val = "0{}".format(v)
            else:
                val = "{}".format(v)
            break
        else:
            print("Please enter a number between {} and {}.".format(min, max))
            continue

    return val

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

    # For each volunteer get the name
    root_dir = adult_volunteer_root_dir
    logging.debug('Volunteer root dir: {}'.format(root_dir))
    for v in volunteers.items():
        v_dir = '{} {}, {}'.format(v[1]['number'], v[1]['last name'], v[1]['first name'])

        # Get the service dates
        for sd in v[1]['service dates']:
            sds = sd.split('-')

            # Get the year
            s_year = os.path.join(v_dir, 'Covid Forms {}'.format(sds[2]))

            # Get the month
            s_month = os.path.join(s_year, calendar.month_name[int(sds[0])])
            s_month_with_root = os.path.join(root_dir, s_month)
            logging.debug("s_month_with_root: {}".format(s_month_with_root))

            # Create or validate the covid form
            fname = '{}.{}.{}.txt'.format(sds[0], sds[1], sds[2])
            fname_with_dir = os.path.join(s_month, fname)
            fname_with_root_dir = os.path.join(root_dir, fname_with_dir)

            logging.debug('file: {}'.format(os.path.join(s_month, fname)))

            if create_form:
                if not os.path.exists(s_month_with_root):
                    os.makedirs(s_month_with_root)
                if not os.path.isfile(fname_with_root_dir):
                    with open(fname_with_root_dir, 'w') as f:
                        f.write('I am cleared to work today!\n')
            else:
                if os.path.isfile(fname_with_root_dir):
                    print('The covid form EXISTS for {}'.format(fname_with_dir))
                else:
                    logging.error('The covid form DOES NOT EXIST for {}'.format(fname_with_dir))


def _show_pdf(pdf_filename):
    logging.debug('_show_pdf({}):'.format(pdf_filename))

    if dry_run:
        return

    if os.name == 'nt':
        # jadfix: try this again
        # fwin = win32gui.GetForegroundWindow()
        _exec_shell_command('start {}'.format(pdf_filename))
        # print("FWIN: {}".format(fwin))
        # time.sleep(10)
        # print("FWIN 2: {}".format(fwin))
        # win32gui.SetActiveWindow(fwin)
        # win32gui.SetForegroundWindow(fwin)
        # print("FWIN 3: {}".format(fwin))
    elif os.name == 'posix':
        _exec_shell_command('open {}'.format(pdf_filename))
    else:
        raise Exception('Unsupported Operating System')

# jadfix: look here
def _create_name_directory_db(root_dir):
    global volunteer_name_dir_db

    logging.debug("_create_name_directory_db({})".format(root_dir))

    for name in os.listdir(root_dir):
        name_with_path = os.path.join(root_dir, name)
        if os.path.isdir(name_with_path):
            first_name = name.split(' ')[2]
            last_name = name.split(' ')[1].rstrip(',')
            vol_num = name.split(' ')[0]
            key = ''.join([first_name, last_name]).lower()
            logging.debug("entry: {} {}".format(key, name_with_path))
            if key in volunteer_name_dir_db.keys():
                volunteer_name_dir_db[key].append(name_with_path)
            else:
                db_entry_list = []
                db_entry_list.append(name_with_path)
                volunteer_name_dir_db[key] = db_entry_list

def _patch_db():
    global volunteer_name_db
    global volunteer_num_db
    global patch_db

    if os.path.exists(patch_db_filename):
        with open(patch_db_filename, 'r') as f:
            patch_db = json.load(f)
            volunteer_name_db |= patch_db
            volunteer_num_db |= patch_db

def _create_db():
    global dirs_to_search
    global volunteer_name_db
    global volunteer_num_db

    logging.debug("_create_db()")

    if os.path.exists(name_db_filename):
        answer = _ask_y_n("The DB files exist do you want to overwrite them? ", default='n')
        if answer == 'n':
            with open(name_db_filename, 'r') as fin:
                volunteer_name_db = json.load(fin)
            with open(num_db_filename, 'r') as fin:
                volunteer_num_db = json.load(fin)
            _patch_db()
            return
    
    print("Creating the databases, this may take a few minutes.....")

    volunteer_name_db = {} 
    volunteer_num_db = {}
   
    for dir in dirs_to_search:
        if dir == None:
            continue

        print("DIR: {}".format(dir))
        for name in os.listdir(dir):
            nm = name.split(',')
            if len(nm) != 2:
                logging.error("The format is invalid for: {}".format(name))
                continue

            last_name = ''.join(nm[0].split(' ')[1:])
            first_name = ''.join(nm[1:])
            vol_num = nm[0].split(' ')[0]
            name_key = (first_name + last_name).replace(' ', '').lower()
            print("NAME KEY: {}".format(name_key))
            print("NUMBER KEY: {}".format(vol_num))
            if name_key in volunteer_name_db:
                entry = volunteer_name_db[name_key]
                print("Previous Entry: {}".format(entry))
                volunteer_name_db[name_key].append(os.path.join(dir, name))
            else:
                entry = []
                entry.append(os.path.join(dir, name))
                volunteer_name_db[name_key] = entry
                print("New Name Entry: {}".format(entry))

            if vol_num in volunteer_num_db:
                entry = volunteer_num_db[vol_num]
                print("Previous Num Entry: {}".format(entry))
                volunteer_num_db[vol_num].append(os.path.join(dir, name))
            else:
                entry = []
                entry.append(os.path.join(dir, name))
                volunteer_num_db[vol_num] = entry
                print("New Num Entry: {}".format(entry))

    _patch_db()

    # print(json.dumps(volunteer_name_db, indent=2))
    # print(json.dumps(volunteer_num_db, indent=2))
    # print(json.dumps(patch_db, indent=2))

    with open(name_db_filename, 'w') as f:
        json.dump(volunteer_name_db, f)
    with open(num_db_filename, 'w') as f:
        json.dump(volunteer_num_db, f)


def _create_directory_db(root_dir):
    global volunteer_dir_db

    # jadfix: there should be a better way to do this
    logging.debug("_build_directory_db({})".format(root_dir))

    if not root_dir:
        return

    rd = os.path.abspath(root_dir)
    for name in os.listdir(rd):
        name_with_path = os.path.join(rd, name)
        logging.debug("name_with_path: {}".format(name_with_path))
        if os.path.isdir(name_with_path):
            vol_num = name.split(' ')[0]
            logging.debug("dir: {} {}".format(vol_num, name_with_path))
            if re.search(r'\d+', vol_num):
                # We only add to the db if it is not in the db and is not a dup
                if vol_num not in volunteer_dir_db and vol_num not in dup_volunteers_db:
                    volunteer_dir_db[vol_num] = name_with_path
                else:
                    logging.debug("Volunteer Number \"{}\" Is a duplicate.".format(vol_num))
                    # If the vol number is in the db, it is the first duplicate. 
                    if vol_num in volunteer_dir_db:
                        dup_volunteers_db[vol_num] = []
                        dup_volunteers_db[vol_num].append(volunteer_dir_db[vol_num])
                        del volunteer_dir_db[vol_num]
                    dup_volunteers_db[vol_num].append(name_with_path)

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

    logging.debug("_get_get_page_filename({})".format(page_number))

    # Get the date to be used in the file name
    if use_previous_date == False:
        month = _ask_value("What is the month on the form", 1, 12, month_on_form)
        if len(month) != 0:
            month_on_form = month

        day = _ask_value("Enter the Day on the form", 0, 31, day_on_form)
        if len(day) != 0:
            day_on_form = day

        year = _ask_value("Enter the Year on the form", 2020, 2030, year_on_form)
        if len(year) != 0:
            year_on_form = year

        answer = _ask_y_n("Do you want to use the same date for every page ", default='n')
        if answer.lower() == 'y':
            use_previous_date = True

    # Get the volunteer number
    v_num = _ask_value("Enter the volunteer number on the form", 0, 99999, None)
    if v_num != '':
        volunteer_number = v_num
    else:
        volunteer_number = '({})'.format(page_number)

    page_filename = '{}_{}_{}-{}.pdf'.format(month_on_form, day_on_form, year_on_form, volunteer_number)
    page_filename = os.path.join(os.path.abspath(forms_dir), page_filename)
  
    return page_filename, volunteer_number


def _split_pdf(file_to_split, create_dir):
    global use_previous_date

    logging.debug('split_pdf({}):'.format(file_to_split))

    _create_db()

    use_previous_date = False
    pdf = PdfFileReader(file_to_split)
    for page in range(pdf.getNumPages()):
        pdf_writer = PdfFileWriter()
        pdf_writer.addPage(pdf.getPage(page))

        tpf = os.path.join(os.path.abspath(forms_dir), tmp_filename)
        logging.debug("Creating a temporary file: {}".format(tpf))
        if not dry_run:
            with open(tpf, 'wb') as out:
                pdf_writer.write(out)

        _show_pdf(tpf)

        single_page_file, volunteer_number = _get_page_filename(page)

        print("Creating file for {}: {}".format(volunteer_number, single_page_file))
        if not os.path.isfile(single_page_file):
            if not dry_run:
                os.rename(os.path.join(forms_dir, tmp_filename), single_page_file)
            if volunteer_number in volunteer_num_db:
                directories = volunteer_num_db[volunteer_number]
                _move_msg(directories, single_page_file)
        else:
            print("The file {} already exists".format(single_page_file))


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
    logging.debug("_execute_move({}, {})".format(os.path.basename(src), dst))

    if dry_run:
        return

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

    filename = os.path.basename(src)
    dst_file = os.path.join(cv_form_dir, filename)
    if not os.path.isfile(dst_file):
        shutil.move(src, cv_form_dir)
    else:
        print("\nThe file {} already exists".format(dst_file))


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

        if re.search(r'(''|[0-1])[0-9]_(''|[0-3])[0-9]_20\d{2}-\d+', src):
            vol_num = re.search(r'-\d+.pdf', src).group().lstrip('-').rstrip('.pdf')
            logging.debug("Vol Number: {}".format(vol_num))

            if vol_num in volunteer_dir_db.keys():
                dst = volunteer_dir_db[vol_num]
                logging.debug("Moving file from: {}".format(src))
                logging.debug("To: {}".format(dst))
                # ans = input("Is this ok y/n [n]?  ")
                # ans.lower()
                # if ans == 'y':
                _execute_move(src, dst)
            else:
                print("The Entry for \"{}\" was not found.".format(os.path.basename(src)))
                if vol_num in dup_volunteers_db:
                    print("DUPLICATES:")
                    for dup in dup_volunteers_db[vol_num]:
                        print("    {}".format(dup))
        else:
            print("\"{}\" was not moved.".format(src))

def _get_file():
    global forms_dir

    logging.debug("_get_file()")

    for fname in os.listdir(forms_dir):
        if os.path.splitext(fname)[len(os.path.splitext(fname)) - 1] != '.pdf':
            continue

        file_with_path = os.path.join(os.path.abspath(forms_dir), fname)
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
    if os.path.splitext(dst)[len(os.path.splitext(dst)) - 1] != '.pdf':
        dst = dst + '.pdf'
    
    logging.debug("The new filename is: {}".format(dst))
    os.rename(src, dst)
    return dst

def split(args):
    """
    Split a pdf file that contains covid attestation forms into multiple files that contain
    a single covid attestation form. The file with the single attestation form is named with a
    name that contains the date on the form and the volunteer number on the form. The file
    will then be moved to tje correct directory location.

    :param args: The parsed input arguments
    :type args: Namespace

    Example: python covid-forms.py split ./scratchroot/covid-forms-Dec-14-15.pdf

    """

    logging.debug("split({})".format(args))

    if args.file:
        file_to_split = args.file
    else:
        file_to_split = _get_file()
        if file_to_split == None:
            print("There is not a valid file to split, please try again.")
            return

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
    global created_file_root
    global volunteer_name_dir_db

    logging.debug("_move_email({}, {}, {})".format(src, clear_date, name))
    key = name.replace(' ', '').lower()
    if key in volunteer_name_dir_db.keys():
        if len(volunteer_name_dir_db[key]) > 1:
            logging.error('There is more than one directory for {}: {}'.format(key, volunteer_name_dir_db[key]))
        else:
            # Rename the file
            dst = volunteer_name_dir_db[key][0]
            vol_num = os.path.basename(dst).split(' ')[0]
            new_file = "{} {}-{}.eml".format(created_file_root, clear_date.replace('/', '_'), vol_num)
            new_file_with_path = os.path.join(os.path.dirname(src), new_file)
            if src != new_file_with_path:
                os.rename(src, new_file_with_path)
            _execute_move(new_file_with_path, dst)
    else:
        logging.error('A directory is not found for {}'.format(key))


def _read_email_eml(filename):
    logging.debug("_read_email_eml({})".format(filename))

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
 
    
def _read_email_msg(filename):
    logging.debug("read_email_msg({})".format(filename))

    msg = MsOxMessage(filename)
    date_time_name = re.findall(r'\d+/\d+/\d+ \d+:\d+:\d+[\r|\n]+[A-Z|a-z| ]+[\r|\n]+', msg.body)
    if len(date_time_name) == 0:
        logging.error("Date, Time and Name was not found for {}!".format(filename))
        return None, None

    date_time_name = date_time_name[0]
    date = re.findall(r'\d+/\d+/\d+', date_time_name)[0].replace('/', '_')
    name = re.findall(r'[\r|\n]+[A-Z|a-z| ]+[\r|\n]+', date_time_name)[0].strip()

    return name, date


def _find_directories(name):
    global dirs_to_search
    directories_found = []

    logging.debug("_find_directory({})".format(name))

    nm = name.split()
    if len(nm) < 2:
        logging.error("Invalid Name: {}".format(name))
        return None

    last_name = nm[len(nm)-1]
    first_name = nm[0]

    # First look for an exact match
    for dir in dirs_to_search:
        if dir:
            directories_found +=  glob.glob(dir + '/**/*{}*{}'.format(last_name, first_name), recursive = True)

    if len(directories_found) == 1:
        print("There is an exact match!")
    else:
        # Look for a last name match
        for dir in dirs_to_search:
            if dir:
                directories_found +=  glob.glob(dir + '/**/*{}*'.format(last_name), recursive = True)

    return directories_found


def find_directories(args):
    logging.debug("find_dir({})".format(args))

    directories = _find_directories("John DeNisco")
    print(directories)


def _move_msg(directories, filename):
    logging.debug("move_msgs(...)")

    if not directories:
        return
    elif len(directories) > 1:
        for dir in directories:
            answer = _ask_y_n("Do you want to move the file {} to {}? ".format(os.path.basename(filename), dir), default='y')
            if answer.lower() == 'y':
                _execute_move(filename, dir)
                return
    else:
        _execute_move(filename, directories[0])


def my_move(args):
    global volunteer_num_db
    global volunteer_name_db
    global patch_db
    
    logging.debug("move_msgs({})".format(args))

    _create_db()

    fd = os.path.abspath(forms_dir)
    for name in os.listdir(fd):
        src = os.path.join(fd, name)
        if not os.path.isfile(src):
            logging.debug("{} is not a file.".format(src))
            continue

        if re.search(r'(''|[0-1])[0-9]_(''|[0-3])[0-9]_20\d{2}-\d+.pdf', name):
            # jadfix: 
            # key = re.search(r'-\d+.pdf', src).group().lstrip('-').rstrip('.pdf')
            key = os.path.splitext(name)[0].split('-')[1]
            db = volunteer_num_db
            logging.debug("Vol Number Key: {}".format(key))
        elif re.search(r'(''|[0-1])[0-9]_(''|[0-3])[0-9]_20\d{2}-(\w| )+.msg', name):
            # jadfix: 
            #key = re.search(r'-(\w| )+.msg', src).group().lstrip('-').rstrip('.msg').replace(' ', '').lower()
            key = os.path.splitext(name)[0].split('-')[1].replace(' ', '')
            db = volunteer_name_db
            logging.debug("Vol Name Key: {}".format(key))
        else:
            logging.error("File is not an email or split page: {}".format(name))
            continue

        if key in db:
            directories = db[key]
            _move_msg(directories, src)
        else:
            logging.debug("Destination not found for {}".format(key))
            if key not in patch_db:
                patch_db[key] = None

    with open(patch_db_filename, 'w') as f:
        json.dump(patch_db, f)


def read_emails(args):
    global volunteer_name_db

    logging.debug("read_emails({})".format(args))

    _create_db()

    fd = os.path.basename(forms_dir)
    for filename in os.listdir(fd):
        file_with_path = os.path.join(fd, filename)
        # jadfix: use os path
        if file_with_path.find('.msg') == -1:
            logging.debug("{} is not an email.".format(file_with_path))
            continue

        name, date = _read_email_msg(file_with_path)
        if name is None:
            print('+++++++++++++++++++++++++++++++++++')
            continue

        new_filename = "{}-{}.msg".format(date, name)
        new_file_with_path = os.path.join(fd, new_filename)
        print('+++++++++++++++++++++++++++++++++++')
        print("Enter Volgistics information for:")
        print("   Date: {}".format(date))
        print("   Name: {}".format(name))
        answer = 'n'
        while answer.lower() != 'y':
            answer = _ask_y_n("Have you entered the information into Volgistics? ", default='y')

        print("Rename the file from {} to {}".format(file_with_path, new_file_with_path))
        if not os.path.exists(new_file_with_path):
            os.rename(file_with_path, new_file_with_path)

        key = name.replace(' ', '').lower()
        if key in volunteer_name_db:
            directories = volunteer_name_db[key]
            _move_msg(directories, new_file_with_path)
        print('+++++++++++++++++++++++++++++++++++')


def mgh_util(args):
    global adult_root_dir
    global junior_root_dir
    global pet_root_dir
    global dry_run

    if args.debug:
        logging.basicConfig(level=logging.DEBUG)

    if args.dry_run:
        dry_run = True

    logging.debug("mgh_util({})".format(args))
    
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
    main_parser.add_argument('--dry-run', '-dr', action='count', help='Print debug output')

    # jadfix: Create should have it's own function
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

    sm_parser = sub_parsers.add_parser('find-dir', help='find dir')
    sm_parser.set_defaults(func=find_directories)

    mv_parser = sub_parsers.add_parser('move', help='Move the files')
    mv_parser.set_defaults(func=my_move)

    main_args = main_parser.parse_args()
    mgh_util(main_args)
