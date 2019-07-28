import requests
import collections
import json
import csv
import os
from getpass import getpass
import re
from config import api_key, proxy_http, proxy_https

#  this name tuple is used to quickly collate time record information
TimeRecords = collections.namedtuple(
    'TimeRecords',
    "contactname,taskname,timetracked,taskid,contactid,projectnumber,projecttitle,tasktotaltimetracked,"
    "tasktimeallocated,id,projectid "
)


def main():
    #  get username, email, password
    windows_username, windows_password, proworkflow_username, proworkflow_password = get_user_details()
    #  setup a session
    session = setup_session(windows_username, windows_password, proworkflow_username, proworkflow_password)
    #  get contacts ids
    contact_ids = get_contact_ids(session)
    #  get time records
    time_records = get_time_records(session, contact_ids)
    #  creates csv ready for use with excel
    with open(os.path.join(os.path.expanduser('~'), 'Downloads', 'time_records.csv'), 'w', newline='') as out:
        file_writer = csv.writer(out)
        file_writer.writerow(['Contact Name', 'Task Name', 'Time Tracked'])
        file_writer.writerows(time_records)


def get_user_details():
    """
    This method gathers their windows username, windows password and proworkflow username/email
    :return: a tuple with the gathered details
    """

    windows_username = input('Enter you windows username or (e[x]it):').lower()
    while re.match(r"[\w]+", windows_username) is None:
        windows_username = input('Invalid: Enter you windows username or (e[x]it):').lower()
        if windows_username == 'x':
            exit()
    windows_password = getpass('Enter your windows password or (e[x]it):')
    while re.match(r"[A-Za-z0-9@#$%^&+=]{5,}", windows_password) is None:
        windows_password = getpass('Invalid: Enter your windows password or (e[x]it):')
        if windows_password == 'x':
            exit()
    proworkflow_username = input('Enter you proworkflow username or (e[x]it):').lower()
    if proworkflow_username == 'x':
        exit()
    proworkflow_password = getpass('Enter your ProWorkflow password or (e[x]it):')
    while re.match(r"[A-Za-z0-9@#$%^&+=]{5,}", proworkflow_password) is None:
        proworkflow_password = getpass('Invalid: Enter your ProWorkflow password or (e[x]it):')
        if proworkflow_password == 'x':
            exit()
    return windows_username, windows_password, proworkflow_username, proworkflow_password


def setup_session(windows_username, windows_password, proworkflow_username, proworkflow_password,
                  session=requests.session()):
    """
    Sets up a session with the proworkflow api so you can make multiple calls on one connection
    :param proworkflow_password: password to authenticate the proworkflow api call
    :param windows_username: username to authenticate the proxy
    :param windows_password: password to authenticate the proxy
    :param proworkflow_username: username to authenticate the proworkflow api call
    :param session: this is a requests session for multiple calls
    :return: returns an authenticated session
    """

    session.auth = (proworkflow_username, proworkflow_password)
    session.proxies = {'http': proxy_http.format(windows_username, windows_password),
                       'https': proxy_https.format(windows_username, windows_password)}
    return session


def get_contact_ids(session):
    """
    gets the contact ids by asking for the users emails, then querying the API to return the contacts ids for each
    which are later used when looking for the time records
    :param session: this is the session that is connected to the proworkflow api
    :return: it returns the contact ids gathered form the api
    """

    emails = input('What contacts would you like information on? (Separate by commas, paste from outlook) '
                   'or (e[x]it):').lower()
    emails = re.findall(r"[\w.\-]+@mbie\.govt\.nz", emails)
    if len(emails) <= 0:
        exit('That input is not valid')
    elif emails == 'x':
        exit()
    emails = [i.strip() for i in emails]
    response = get_response_data(session, 'https://api.proworkflow.net/contacts?fields=email,name&apikey={}'
                                 .format(api_key))
    data = json.loads(response.text)
    data = data['contacts']
    contacts_ids = [str(i['id']) for i in data if i['email'].lower() in emails]
    return contacts_ids


def get_time_records(session, contact_ids):
    """
    this gathers time records using the API for each contact id provided
    :param session: this is the session connected the the proworkflow API
    :param contact_ids: these ids are used to query the API for time record information
    :return: returns time records with the value converted to a double for easy manipulation in excel
    """

    time_from = input('Please enter the date from which you wish to start tracking(dd-mm-yyyy) or (e[x]it):')
    while re.match(r"[0-9]{2}-[0-9]{2}-[0-9]{4}", time_from) is None:
        time_from = input('Invalid: '
                          'Please enter the date from which you wish to start tracking(dd-mm-yyyy) or (e[x]it):')
        if time_from == 'x':
            exit()
    time_from = time_from.split('-')
    time_from.reverse()
    time_from = '-'.join(time_from)
    time_to = input('Please enter the date to which you wish to end tracking(dd-mm-yyyy) or (e[x]it):')
    while re.match(r"[0-9]{2}-[0-9]{2}-[0-9]{4}", time_to) is None:
        time_to = input('Invalid: Please enter the date to which you wish to end tracking(dd-mm-yyyy) or (e[x]it):')
        if time_to == 'x':
            exit()
    time_to = time_to.split('-')
    time_to.reverse()
    time_to = '-'.join(time_to)
    #  contact_ids = [str(i) for i in contact_ids]
    response = get_response_data(session, 'https://api.proworkflow.net/'
                                          'time?trackedfrom={}&trackedto={}&'
                                          'fields=id,task,project,tasktimetotals,timetracked,contact&contacts={}'
                                          '&apikey={}'
                                 .format(time_from, time_to, ','.join(contact_ids), api_key))
    data = json.loads(response.text)
    data = data['timerecords']
    time_records = [
        TimeRecords(**td)
        for td in data
    ]
    time_records = [(i.contactname, i.taskname, i.timetracked/1440) for i in time_records]
    return time_records


def get_response_data(session, url):
    """
    This method is used to perform queries on the session. Having a separate method mean you only have to perform error
    handling in one place.
    :param session:
    :param url:
    :return: returns json data
    """

    try:
        response = session.get(url)
        response.raise_for_status()
        return response
    except Exception as e:
        print(e)
        exit()


if __name__ == '__main__':
    main()
