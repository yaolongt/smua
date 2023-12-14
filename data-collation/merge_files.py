import copy
from datetime import datetime as dt
import json
import pandas as pd
import os
import warnings

warnings.simplefilter("ignore")

# Define the headers to read in
session_headers = ['Dept', 'Course Type', 'Sch #', 'Related Schedule #', 'Session #',
                     'Session Date', 'Session Day', 'S-Time', 'E-Time', 'Venue', 'Lecturer']
schedule_headers = ['Course Type', 'Sch #', 'Schedule Audience', 'Client Name',
                          'Course RunID', 'Course Title', 'Sch S-Date', 'Sch E-Date', 'Sch Status', 'Enr Pax']
enrolment_headers = ["Schedule #", "# Registered"]

# Define new column names for new file
new_cols = ['Pillar', 'Course No.', 'Course Title', 'Status', 'Course Run ID', 'Mode of Delivery', \
            'Type of Runs (Public or Corporate)', 'Start Date', 'End Date', 'Session Date & Time', \
            'Session Venue', 'Location by Date', 'Total no. of sessions', 'Registered Pax', 'Enrolled Pax',\
            'Total Pax', 'Venue Category', 'Last Updated']


def get_data_from_file():
    """
    To read data from a file called `data.json`.
    Data consists of `days` and `schools` that is being used in the code.
    """
    with open("./data.json", 'r') as file:
        data = json.load(file)
        schools = set(data['schools'])
        days = data['days']

        return schools, days


def convert_to_dict(session, schedule, enroll):
    """
    Convert file from dataframe to JSON key-value pair
    """
    session_map = session.to_dict()
    schedule_map = schedule.set_index("Sch #").T.to_dict()
    enroll_map = enroll.set_index("Schedule #").T.to_dict()

    return session_map, schedule_map, enroll_map


def read_files():
    """
    Read excel files and replace empty values with '-'
    """
    session = pd.read_excel("gvSession.xlsx", usecols=session_headers).fillna("-")
    schedule = pd.read_excel("Manage Schedule.xlsx", usecols=schedule_headers).fillna("-")
    enroll = pd.read_excel("Enrolment Summary.xlsx", usecols=enrolment_headers).fillna("-")

    return session, schedule, enroll


def get_combined_values(session):
    """
    Combine values of "Session Date + Session Time" and "Session + Venue"
    """

    # combine all the values to form "Session Date + Session Time"
    session_datetime = session['Session Date'].astype('datetime64[ns]').dt.strftime('%Y-%m-%d') \
        + ' ' + session['Course Type'].str[0] + session['Session #'].astype(str) \
        + ' : ' + session['Session Day'] + ' ' \
        + session['S-Time'] + ' to ' + session['E-Time']

    # combine all the values to form "Session + Venue"
    session_venue = session['Session Date'].astype('datetime64[ns]').dt.strftime('%Y-%m-%d') \
        + ' ' + session['Course Type'].str[0] \
        + session['Session #'].astype(str) + ' - Venue: ' + session['Venue']

    return session_datetime, session_venue


def get_course_audience(schedule_map):
    """
    This function gets course audiences, with a certain formatting.
    """
    course_audience_map = {}
    for key, value in schedule_map.items():
        schedule_audience = value['Schedule Audience']

        if schedule_audience == '-':
            course_audience_map[key] = '-'
        else:
            client_name = value['Client Name']
            if client_name == '-':
                course_audience_map[key] = schedule_audience
            else:
                course_audience_map[key] = value['Schedule Audience'] + " : " + value['Client Name']

    return course_audience_map


def map_sessions(session, session_map):
    """
    Combine all the values of Session Date Time and Session Venue to a 'key'.
    Get the pillar of the particular Schedule using the Session.
    The 'key' will be the 'Sch #'.
    """
    session_datetime, session_venue = get_combined_values(session)

    name_dict = {'Finance & Technology': 'FIT',
                 'Human Capital, Management & Leadership': 'HCML',
                 'Business Management': 'BM',
                 'Services, Operations and Business Improvement': 'SOBI'}

    sessions_details = {
        key: {
            "pillar": None,
            "assessment": {
                "datetime": [],
                "venue": []
            },
            "normal": {
                "datetime": [],
                "venue": []
            }
        } for key in session_map["Sch #"].values()
    }

    for i in range(len(session)):
        schedule_no = session_map['Sch #'][i]

        if not sessions_details[schedule_no]["pillar"]:
            dept = name_dict.get(session_map['Dept'][i])
            if not dept:
                dept = "No dept"

            sessions_details[schedule_no]["pillar"] = dept

        if session_map['Course Type'][i] == 'Assessment':
            schedule_no = session_map['Related Schedule #'][i]

            # Check if the 'Related Schedule #' is in inside the map as key
            if schedule_no not in sessions_details:
                continue
            sessions_details[schedule_no]["assessment"]["datetime"].append(session_datetime[i])
            sessions_details[schedule_no]["assessment"]["venue"].append(session_venue[i])
        else:
            sessions_details[schedule_no]["normal"]["datetime"].append(session_datetime[i])
            sessions_details[schedule_no]["normal"]["venue"].append(session_venue[i])

    return sessions_details


def format_location_by_date(sorted_venue):
    """
    This function groups a location by date.
    The group will use the first date by each location.

    Assumption made: Only 1 location per day for each session.
    """
    session_venues_map = {}
    for session in sorted_venue:
        parts = session.split("Venue:")

        session_date = parts[0].split(" ")[0] 

        if session_date in session_venues_map: 
            continue

        session_venue = parts[1][1:]
        session_venues_map[session_date] = session_venue

    res = ""
    for k, v in session_venues_map.items():
        res += f'{k} \n{v}\n\n'

    return res

def find_venue_type(venue, schools):
    """
    Get the venue type based on session's venue
    """
    if 'Venue: -' in venue or 'Venue: Cancelled' in venue:
        return venue.split("Venue:")[0] + '-'
    elif 'Online' in venue:
        return venue.split("Venue:")[0] + 'Online'
    else:
        for building in schools:
            if building in venue:
                return venue.split("Venue:")[0] + 'Onsite'
        return venue.split("Venue:")[0] + 'Offsite'


def change_venue_names(venue):
    """
    Replace venue name's shortcut
    """
    venue = venue.replace("SR", " Seminar Room")
    venue = venue.replace("CR", " Classroom")
    venue = venue.replace(" SMU ", " ")
    return venue


def add_total_pax(registered_pax, enr_pax):
    """
    Calculate the total pax based on Registered and Enrolled Pax
    """
    if registered_pax == '-':
        if enr_pax == '-':
            return 0
        else:
            return enr_pax
    else:
        return enr_pax + registered_pax


def structure_data(schedule_map, sessions_details, enroll_map, audience_map, schools):
    """
    This function is to structure the data accord to the output.
    Do note that there are quite a number of data manipulation to get the desired output.
    """
    res = []

    for key, value in schedule_map.items():
        if value['Course Type'] == 'Assessment':
            continue

        if not sessions_details.get(key):
            continue

        # Extract values
        pillar = sessions_details.get(key).get('pillar')

        if not pillar:
            continue

        title = value['Course Title']
        status = value['Sch Status']
        runid = value['Course RunID']

        # Join the Normal Session datetime and Assessment Session datetime in order
        datetime_data = copy.deepcopy(sorted(sessions_details[key]["normal"]["datetime"]))
        datetime_data.extend(sorted(sessions_details[key]["assessment"]["datetime"]))
        datetime = " \n".join(datetime_data)

        # Join the Normal Session venue and Assessment Session venue in order
        venue_data = copy.deepcopy(sorted(sessions_details[key]["normal"]["venue"]))
        venue_data.extend(sorted(sessions_details[key]["assessment"]["venue"]))

        sorted_venue = []
        category_venue = []

        for venue in venue_data:
            if venue == '-':
                continue
            category_venue.append(find_venue_type(venue, schools))    # Label venue type
            sorted_venue.append(change_venue_names(venue))   # Format short forms

        location_by_date = format_location_by_date(sorted_venue)
        session_venue = " \n".join(sorted_venue)
        delivery_mode = "F2F" if "Online" not in session_venue else "Online"
        course_audience = audience_map[key]
        start_date = value['Sch S-Date'].strftime('%Y-%m-%d')
        end_date = value['Sch E-Date'].strftime('%Y-%m-%d')
        no_sessions = f'No. of sessions: {len(sessions_details[key]["normal"]["datetime"])} \nNo. of assessments: {len(sessions_details[key]["assessment"]["datetime"])}'
        enrolled_pax = value['Enr Pax']
        registered_pax = enroll_map[key]['# Registered'] if key in enroll_map else '-'
        total_pax = add_total_pax(registered_pax, enrolled_pax)

        data = [pillar, key, title, status, runid, delivery_mode,\
                course_audience, start_date, end_date, datetime, session_venue,\
                location_by_date, no_sessions, registered_pax, enrolled_pax, total_pax,\
                " \n".join(category_venue), f'Last Updated: -']

        res.append(data)

    return res

def find_course_more_than_6days(data, days):
    res = []
    for i in range(len(data)):
        start_date = dt.strptime(data[i][7], '%Y-%m-%d').date()
        end_date = dt.strptime(data[i][8], '%Y-%m-%d').date()
        
        # if the course start and end dates are more than 6 days apart, add into list
        if (end_date - start_date).days > days:
            res.append(data[i])

    return res

def format_cells(workbook, worksheet):
    normal_text = workbook.add_format({'text_wrap': True})
    bold_text = workbook.add_format({'text_wrap': True, 'bold': True, 'font_size': 15})
    header_format = workbook.add_format({'bold': True, 'fg_color': "#ffcccc", 'border': 1, 'font_size': 15})


    # To color the header column and bold it
    for colno, value in enumerate(data_df.columns.values):
        worksheet.write(0, colno, value, header_format)

    worksheet.set_column('A:A', 20, normal_text)
    worksheet.set_column('B:C', 40, normal_text)
    worksheet.set_column('D:G', 20, normal_text)
    worksheet.set_column('H:I', 20, bold_text)
    worksheet.set_column('J:L', 40, bold_text)
    worksheet.set_column('K:K', 80, bold_text)
    worksheet.set_column('M:M', 20, normal_text)
    worksheet.set_column('N:P', 20, bold_text)
    worksheet.set_column('Q:R', 20, normal_text)


if __name__ == "__main__":
    if "gvSession.xlsx" not in os.listdir() \
        or "Manage Schedule.xlsx" not in os.listdir() \
        or "Enrolment Summary.xlsx" not in os.listdir():
        exit("Files are missing!")
    
    session, schedule, enroll = read_files()
    schools, days = get_data_from_file()

    session_map, schedule_map, enroll_map = convert_to_dict(session, schedule, enroll)
    sessions_details = map_sessions(session, session_map)
    audience_map = get_course_audience(schedule_map)

    data = structure_data(schedule_map, sessions_details, enroll_map, audience_map, schools)

    current_datetime = dt.now().strftime("%Y%m%d_%H%M")
    filename = f'CDL_{current_datetime}.xlsx'

    data_df = pd.DataFrame(data, columns=new_cols)
    data_df = data_df.sort_values(by=['Start Date', 'Course No.'])

    writer = pd.ExcelWriter(filename, engine='xlsxwriter')

    data_df.to_excel(writer, sheet_name='Sheet1', index=False)

    workbook = writer.book

    # Raw CDL Data in Sheet1
    worksheet = writer.sheets['Sheet1']
    format_cells(workbook, worksheet)

    # Data where start and end date is more than 6 days
    long_period = find_course_more_than_6days(data, days)
    long_period_df = pd.DataFrame(long_period, columns=new_cols).sort_values(by=['Start Date', 'End Date'])

    long_period_df.to_excel(writer, sheet_name='Course > 6 days', index=False)

    worksheet = writer.sheets['Course > 6 days']
    format_cells(workbook, worksheet)

    writer.close()
    
    print(f"File compile successful. File name: {filename}\n")
    exit("Finish execution.")
