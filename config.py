from dateutil.rrule import SA, SU, MO, TU, WE, TH, FR

# Timesheet Information
TIMESHEET = 'C:/Users/hhoyos/ownCloud/Timesheets/Zeiterfassung_Horacio Hoyos Rodriguez_2017.xlsx'
NAME = 'Horacio Hoyos Rodriguez'
DEPARTMENT = 'ISSE'
INIT = False

# Toggl API Token
API_TOKEN = 'fbf549bd1a1102e76b56d9a61c36882d'
EMAIL = 'arcanefoam@gmail.com'

# Number Precision
PRECISION = 2

# Target Configurations
WORKING_SCHEDULE_PER_DAY = {'Mo': '08:00-14:00', 'Di': '08:00-14:00', 'Mi': '08:00-14:00',
                            'Do': '08:00-14:00', 'Fr': '08:00-14:00'}
TOLERANCE_PERCENTAGE = 0.1

# Timezone
TIMEZONE = 'Europe/Vienna'

# Projects/Clients
# Add/change to other clients/projects that you want to track for your hours. Projects and clients
# are organized by workspace. The project id can be found from the URL of the project information page,
# e.g.:
# https://toggl.com/app/projects/1584496/edit/27175904
# https://toggl.com/app/projects/{workspace_id}/edit/{project_id}
WORK_SPACES = {'1584496': {'clients': ['JKU', 'Me'],
                           'projects': ['27175904', '24160610']}}
