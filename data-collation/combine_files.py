import pandas as pd
from datetime import datetime as dt
import os

# Define the headers to read in
gvSession_headers = ['Dept', 'Course Type', 'Sch #', 'Related Schedule #', 'Session #',
                     'Session Date', 'Session Day', 'S-Time', 'E-Time', 'Venue', 'Lecturer']
manageSchedule_headers = ['Course Type', 'Sch #', 'Schedule Audience', 'Client Name',
                          'Course RunID', 'Course Title', 'Sch S-Date', 'Sch E-Date', 'Sch Status', 'Enr Pax']
enrolment_summary_headers = ["Schedule #", "# Registered"]

# Define new column names for new file
new_cols = ['Pillar', 'Course No.', 'Course Title', 'Status', 'Course Run ID', 'Mode of Delivery', \
            'Type of Runs (Public or Corporate)', 'Start Date', 'End Date', 'Session Date & Time', \
            'Session Venue', 'Location by Date', 'Total no. of sessions', 'Registered Pax', 'Enrolled Pax',\
            'Total Pax', 'Venue Category', 'Last Updated']

schools = set()

gvSession, manageSchedule, enrolmentSummary = None, None, None
gvSessionMap, manaageScheduleMap, enrolmentSummaryMap = None, None, None


def get_school_buildings():
    """To read school buildings from a text file `schools.txt` instead of hard-coding here.
        Insert values into a dictionary.
    """
    with open("schools.txt") as f:
        for building in f:
            schools.add(building.strip())


"""
    Convert into a dict type
"""
def get_file_dict():
    global gvSessionMap, manageScheduleMap, enrolmentSummaryMap
    gvSessionMap = gvSession.to_dict()
    manageScheduleMap = manageSchedule.to_dict()
    enrolmentSummaryMap = enrolmentSummary.set_index("Schedule #").T.to_dict()


"""
    Read Excel files
"""
def read_files():
    global gvSession, manageSchedule, enrolmentSummary
    if "gvSession.xlsx" not in os.listdir() or "Manage Schedule.xlsx" not in os.listdir() or "Enrolment Summary.xlsx" not in os.listdir():
        return False

    gvSession = pd.read_excel("gvSession.xlsx", usecols=gvSession_headers)
    manageSchedule = pd.read_excel("Manage Schedule.xlsx", usecols=manageSchedule_headers)
    enrolmentSummary = pd.read_excel("Enrolment Summary.xlsx", usecols=enrolment_summary_headers)
    gvSession = gvSession.fillna("-")
    manageSchedule = manageSchedule.fillna("-")
    enrolmentSummary = enrolmentSummary .fillna("-")

    get_file_dict()

    return True


"""
    Function to get all the combined values like "Session Date + Session Time" and "Session + Venue"
"""
def get_combined_values():
    # combine all the values to form "Session Date + Session Time"
    session_datetime = gvSession['Session Date'].astype('datetime64[ns]').dt.strftime('%Y-%m-%d') + ' ' + gvSession['Course Type'].str[0] + gvSession['Session #'].astype(str) \
        + ' : ' + gvSession['Session Day'] + ' ' + \
        gvSession['S-Time'] + ' to ' + gvSession['E-Time']

    # combine all the values to form "Session + Venue"
    session_venue = gvSession['Session Date'].astype('datetime64[ns]').dt.strftime('%Y-%m-%d') + ' ' + gvSession['Course Type'].str[0] + \
        gvSession['Session #'].astype(str) + ' - Venue: ' + gvSession['Venue']

    return session_datetime, session_venue


"""
    Function to map sessions information together
"""
def map_sessions():
    session_datetime, session_venue = get_combined_values()

    session_datetime_map = {v: [] for v in gvSessionMap['Sch #'].values()}
    session_venue_map = {v: [] for v in gvSessionMap['Sch #'].values()}

    for i in range(len(gvSession)):
        course_no = gvSessionMap['Sch #'][i]
        session_datetime_map[course_no].append(session_datetime[i])
        session_venue_map[course_no].append(session_venue[i])
    
    return session_datetime_map, session_venue_map


"""
    Function to combine course audience
"""
def get_course_audience():
    course_audience_map = {}
    for i in range(len(manageSchedule)):
        course_no = manageScheduleMap['Sch #'][i]
        schedule_audience =  manageScheduleMap['Schedule Audience'][i]

        if schedule_audience == '-':
            course_audience_map[course_no] = '-'
        else:
            client_name = manageScheduleMap['Client Name'][i]
            if client_name == '-':
                course_audience_map[course_no] = schedule_audience
            else:
                course_audience_map[course_no] = manageScheduleMap['Schedule Audience'][i] + " : " + manageScheduleMap['Client Name'][i]
    
    return course_audience_map


"""
    Get course pillar column according to the Course No.
"""
def get_course_pillar():
    course_pillar_map = {}
    for i in range(len(gvSession)):
        course_pillar_map[gvSessionMap['Sch #'][i]] = gvSessionMap['Dept'][i]
    
    return course_pillar_map


"""
    Helper function to modify the venue type
"""
def find_venue_type(venue):
    if 'Venue: -' in venue or 'Venue: Cancelled' in venue:
        return venue.split("Venue:")[0] + '-'
    elif 'Online' in venue:
        return venue.split("Venue:")[0] + 'Online'
    else:
        for building in schools:
            if building in venue:
                return venue.split("Venue:")[0] + 'Onsite'
        return venue.split("Venue:")[0] + 'Offsite'


"""
    Helper function to modify the SR/CR to Seminar Room/Classroom:
"""
def change_venue_short_form(venue):
    venue = venue.replace("SR", " Seminar Room")
    venue = venue.replace("CR", " Classroom")
    venue = venue.replace(" SMU ", " ")
    return venue

def add_total_pax(registered_pax, enr_pax):
    if registered_pax == '-':
        if enr_pax == '-':
            return 0
        else:
            return enr_pax
    else:
        return enr_pax + registered_pax

def session_venues_format(sorted_venue):
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

"""
    Format the data in order so that it can be exported properly    
"""
def structure_data_without_assessments(session_datetime_map, session_venue_map, course_audience_map, course_pillar_map):
    new_data = []
    new_data_map = {}
    name_dict = {'Finance & Technology': 'FIT',
                 'Human Capital, Management & Leadership': 'HCML',
                 'Business Management': 'BM',
                 'Services, Operations and Business Improvement': 'SOBI'}
    
    for i in range(len(manageSchedule)):
        # Skip "Assessment" type
        if manageScheduleMap['Course Type'][i] == 'Assessment':
            continue

        course_no = manageScheduleMap['Sch #'][i]

        # if there are no sessions for the module
        if not course_pillar_map.get(course_no):
            continue

        # Get all the values that are corresponding to this current 'Course No'
        pillar = name_dict.get(course_pillar_map[course_no]) if course_pillar_map[course_no] in name_dict else "No dept"
        course_title = manageScheduleMap['Course Title'][i]
        status = manageScheduleMap['Sch Status'][i]
        course_runid = manageScheduleMap['Course RunID'][i]
        session_datetime = " \n".join(sorted(session_datetime_map[course_no]))
                
        sorted_venue = []
        category_venue = []
        
        
        for venue in session_venue_map[course_no]:
            if venue == '-':
                continue
            category_venue.append(find_venue_type(venue))
            venue = change_venue_short_form(venue)
            sorted_venue.append(venue)

        category_venue.sort()
        sorted_venue.sort()
        formatted_session_venue = session_venues_format(sorted_venue)
        session_venue = " \n".join(sorted_venue)
        mode_delivery = "F2F" if "Online" not in formatted_session_venue else "Online"
        course_audience = course_audience_map[course_no]
        start_date = manageScheduleMap['Sch S-Date'][i].strftime('%Y-%m-%d')
        end_date = manageScheduleMap['Sch E-Date'][i].strftime('%Y-%m-%d')
        no_sessions = "No. of sessions: " + str(len(session_datetime_map[course_no]))
        enr_pax = manageScheduleMap['Enr Pax'][i]

        registered_pax = enrolmentSummaryMap[course_no]['# Registered'] if course_no in enrolmentSummaryMap else '-'

        total_pax = add_total_pax(registered_pax, enr_pax)
        
        temp = [pillar, course_no, course_title, status, course_runid, mode_delivery, course_audience,
                   start_date, end_date, session_datetime, session_venue, formatted_session_venue, no_sessions, registered_pax, 
                   enr_pax, total_pax, " \n".join(category_venue), f'Last Updated: -']
        new_data.append(temp)
        new_data_map[course_no] = temp
    
    return new_data, new_data_map


"""
    Add Assessments to assessment_map
    This function helps to map all the relevant assessments to relevant course code.
    A module may have multiple assessments.
"""
def add_assessment_to_map(new_data_map, session_datetime_map, session_venue_map):
    assessment_map = {}

    for i in range(len(gvSession)):
        if gvSessionMap['Course Type'][i] == 'Assessment':
            course_no = gvSessionMap['Related Schedule #'][i]
            if course_no in new_data_map:
                session_date = dt.strftime(gvSessionMap['Session Date'][i], '%Y-%m-%d')
                assessment_datetime = f"{session_date} A{gvSessionMap['Session #'][i]} : {gvSessionMap['Session Day'][i]} {gvSessionMap['S-Time'][i]} to {gvSessionMap['E-Time'][i]}"
                venue = change_venue_short_form(gvSessionMap['Venue'][i])
                assessment_venue = f"{session_date} A{gvSessionMap['Session #'][i]} - Venue: {venue}"

                course_assessments = assessment_map.get(course_no, [])
                
                """
                    Structure the assessment data by "Datetime", "Venue", "Type of Venue" for easier access
                """
                course_assessments.append([assessment_datetime, assessment_venue, find_venue_type(assessment_venue)])
                assessment_map[course_no] = course_assessments
    
    return assessment_map


"""
  Add relevant assessments to the structured data
"""
def add_assessment_to_structured_data(new_data_map, assessment_map):
    for course_no, data in assessment_map.items():
        no_of_assessments = len(data)
        new_data_map[course_no][12] += f" \nNo. of assessments: {no_of_assessments}"
        for list in data:
            session_datetime = list[0]
            session_venue = change_venue_short_form(list[1])
            session_type = list[2]
            new_data_map[course_no][9] += f" \n{session_datetime}"
            new_data_map[course_no][10] += f" \n{session_venue}"
            new_data_map[course_no][11] += f"\n{session_venue}"
            new_data_map[course_no][16] += f" \n{session_type}"


"""
    Starting point of the Python code
"""
if __name__ == "__main__":
    if not read_files():
        exit("Files are missing!")
    
    get_school_buildings()
    session_datetime_map, session_venue_map = map_sessions()
    course_audience_map = get_course_audience()
    course_pillar_map = get_course_pillar()

    new_data, new_data_map = structure_data_without_assessments(session_datetime_map, session_venue_map, course_audience_map, course_pillar_map)
    assessment_map = add_assessment_to_map(new_data_map, session_datetime_map, session_venue_map)
    add_assessment_to_structured_data(new_data_map, assessment_map)

    current_datetime = dt.now().strftime("%Y%m%d_%H%M")
    new_data = pd.DataFrame(new_data, columns=new_cols)

    new_data = new_data.sort_values(by=['Start Date', 'Course No.'])
    filename = f'CDL_{current_datetime}.xlsx'

    new_data.to_excel(filename, sheet_name='Sheet1', engine='openpyxl')
    writer = pd.ExcelWriter(filename, engine="xlsxwriter")
    new_data.to_excel(writer, sheet_name='Sheet1', index=False)

    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]

    normal_text = workbook.add_format({'text_wrap': True})
    bold_text = workbook.add_format({'text_wrap': True, 'bold': True, 'font_size': 15})
    header_format = workbook.add_format({'bold': True, 'fg_color': "#ffcccc", 'border': 1, 'font_size': 15})


    # To color the header column and bold it
    for colno, value in enumerate(new_data.columns.values):
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
