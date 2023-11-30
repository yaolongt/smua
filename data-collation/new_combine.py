import copy
from datetime import datetime as dt
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

schools = set()

def get_school_buildings():
    """
    To read school buildings from a text file `schools.txt`.
    Insert values into a dictionary.
    """
    with open("schools.txt") as f:
        for building in f:
            schools.add(building.strip())


def convert_to_dict(session, schedule, enroll):
    """
    Convert file from dataframe to JSON key-value pair
    """
    session_map = session.to_dict()
    schedule_map = schedule.to_dict()
    enroll_map = enroll.set_index("Schedule #").T.to_dict()

    return session_map, schedule_map, enroll_map


def read_files():
    """
    Read excel files and replace empty values with '-'
    """
    session = pd.read_excel("gvSession.xlsx", usecols=session_headers)
    schedule = pd.read_excel("Manage Schedule.xlsx", usecols=schedule_headers)
    enroll = pd.read_excel("Enrolment Summary.xlsx", usecols=enrolment_headers)

    # fill empty values with '-'
    session = session.fillna("-")
    schedule = schedule.fillna("-")
    enroll = enroll.fillna("-")

    get_school_buildings()

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


def get_course_audience(schedule, schedule_map):
    """
    This function gets course audiences, with a certain formatting.
    """
    course_audience_map = {}
    for i in range(len(schedule)):
        course_no = schedule_map['Sch #'][i]
        schedule_audience =  schedule_map['Schedule Audience'][i]

        if schedule_audience == '-':
            course_audience_map[course_no] = '-'
        else:
            client_name = schedule_map['Client Name'][i]
            if client_name == '-':
                course_audience_map[course_no] = schedule_audience
            else:
                course_audience_map[course_no] = schedule_map['Schedule Audience'][i] + " : " + schedule_map['Client Name'][i]
    
    return course_audience_map


def get_course_pillar(session, session_map):
    """
    This function gets all course pillars related to the course.
    """
    name_dict = {'Finance & Technology': 'FIT',
                 'Human Capital, Management & Leadership': 'HCML',
                 'Business Management': 'BM',
                 'Services, Operations and Business Improvement': 'SOBI'}
    pillar_map = {}

    for i in range(len(session)):
        dept = name_dict.get(session_map['Dept'][i])
        if not dept:
            dept = "No dept"

        pillar_map[session_map['Sch #'][i]] = dept

    return pillar_map


def map_sessions(session, session_map):
    """
    Combine all the values of Session Date Time and Session Venue to a 'key'.
    The 'key' will be the 'Sch #'.
    """
    session_datetime, session_venue = get_combined_values(session)

    # Set 'Schedule #' to be the key to link all the session together
    # `[[<Session>],[<Assessment>]]` is to keep the Assessment type and the normal session differentiated.
    session_datetime_map = {key: [[], []] for key in session_map["Sch #"].values()}
    session_venue_map = {key: [[], []] for key in session_map["Sch #"].values()}

    for i in range(len(session)):
        schedule_no = session_map['Sch #'][i]

        # Add type of 'Assessment' into the map
        if session_map['Course Type'][i] == 'Assessment': 
            schedule_no = session_map['Related Schedule #'][i]

            # Check if the 'Related Schedule #' is in inside the map as key
            if schedule_no not in session_datetime_map:
                continue
            
            session_datetime_map[schedule_no][1].append(session_datetime[i])
            session_venue_map[schedule_no][1].append(session_venue[i])
        else:
            session_datetime_map[schedule_no][0].append(session_datetime[i])
            session_venue_map[schedule_no][0].append(session_venue[i])

    return session_datetime_map, session_venue_map


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

def find_venue_type(venue):
    """
    Get the venue type based on session's venue
    """
    # print(schools)
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


def structure_data(schedule, schedule_map, session_datetime_map, session_venue_map, enroll_map, audience_map, pillar_map):
    """
    This function is to structure the data accord to the output.
    Do note that there are quite a number of data manipulation to get the desired output.
    """
    res = []

    for i in range(len(schedule)):
        if schedule_map['Course Type'][i] == 'Assessment':
            continue

        # Extract values
        schedule_no = schedule_map['Sch #'][i]
        pillar = pillar_map.get(schedule_no)

        if not pillar:
            continue

        title = schedule_map['Course Title'][i]
        status = schedule_map['Sch Status'][i]
        runid = schedule_map['Course RunID'][i]

        # Join the Normal Session datetime and Assessment Session datetime in order
        datetime_data = copy.deepcopy(sorted(session_datetime_map[schedule_no][0]))
        datetime_data.extend(sorted(session_datetime_map[schedule_no][1]))
        datetime = " \n".join(datetime_data)

        # Join the Normal Session venue and Assessment Session venue in order
        venue_data = copy.deepcopy(sorted(session_venue_map[schedule_no][0]))
        venue_data.extend(sorted(session_venue_map[schedule_no][1]))

        sorted_venue = []
        category_venue = []

        for venue in venue_data:
            if venue == '-':
                continue
            category_venue.append(find_venue_type(venue))    # Label venue type
            sorted_venue.append(change_venue_names(venue))   # Format short forms

        location_by_date = format_location_by_date(sorted_venue)
        session_venue = " \n".join(sorted_venue)
        delivery_mode = "F2F" if "Online" not in session_venue else "Online"
        course_audience = audience_map[schedule_no]
        start_date = schedule_map['Sch S-Date'][i].strftime('%Y-%m-%d')
        end_date = schedule_map['Sch E-Date'][i].strftime('%Y-%m-%d')
        no_sessions = f"No. of sessions: {len(session_venue_map[schedule_no][0])} \nNo. of assessments: {len(session_venue_map[schedule_no][1])}"
        enrolled_pax = schedule_map['Enr Pax'][i]
        registered_pax = enroll_map[schedule_no]['# Registered'] if schedule_no in enroll_map else '-'
        total_pax = add_total_pax(registered_pax, enrolled_pax)

        data = [pillar, schedule_no, title, status, runid, delivery_mode,\
                course_audience, start_date, end_date, datetime, session_venue,\
                location_by_date, no_sessions, registered_pax, enrolled_pax, total_pax,\
                " \n".join(category_venue), f'Last Updated: -']

        res.append(data)

    return res

if __name__ == "__main__":
    if "gvSession.xlsx" not in os.listdir()\
        or "Manage Schedule.xlsx" not in os.listdir()\
        or "Enrolment Summary.xlsx" not in os.listdir():
        exit("Files are missing!")
    
    
    session, schedule, enroll = read_files()
    session_map, schedule_map, enroll_map = convert_to_dict(session, schedule, enroll)
    session_datetime_map, session_venue_map = map_sessions(session, session_map)
    audience_map = get_course_audience(schedule, schedule_map)
    pillar_map = get_course_pillar(session, session_map)

    data = structure_data(schedule, schedule_map, session_datetime_map, session_venue_map, enroll_map, audience_map, pillar_map)

    current_datetime = dt.now().strftime("%Y%m%d_%H%M")
    filename = f'CDL_{current_datetime}.xlsx'

    data_df = pd.DataFrame(data, columns=new_cols)
    data_df = data_df.sort_values(by=['Start Date', 'Course No.'])

    writer = pd.ExcelWriter(filename, engine='xlsxwriter')

    data_df.to_excel(writer, sheet_name='Sheet1', index=False)

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

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

    writer.close()
    print(f"File compile successful. File name: {filename}\n")
    exit("Finish execution.")
