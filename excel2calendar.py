import os
import csv
import StringIO
import urllib

import jinja2
import webapp2

from apiclient import discovery
from oauth2client import appengine
from oauth2client import client

JINJA_ENVIRONMENT = jinja2.Environment(
    loader=jinja2.FileSystemLoader(os.path.dirname(__file__)),
    extensions=['jinja2.ext.autoescape'],
    autoescape=True)

# CLIENT_SECRETS is name of a file containing the OAuth 2.0 information for this
# application, including client_id and client_secret. You can see the Client ID
# and Client secret on the APIs page in the Cloud Console:
# <https://cloud.google.com/console#/project/729375114005/apiui>
CLIENT_SECRETS = os.path.join(os.path.dirname(__file__), 'client_secrets.json')

# Helpful message to display in the browser if the CLIENT_SECRETS file
# is missing.
MISSING_CLIENT_SECRETS_MESSAGE = """
<h1>Warning: Please configure OAuth 2.0</h1>
<p>
To make this sample run you will need to populate the client_secrets.json file
found at:
</p>
<p>
<code>%s</code>.
</p>
<p>with information found on the <a
href="https://code.google.com/apis/console">APIs Console</a>.
</p>
""" % CLIENT_SECRETS

service = discovery.build('calendar', 'v3')
decorator = appengine.oauth2decorator_from_clientsecrets(
    CLIENT_SECRETS,
    scope=[
        'https://www.googleapis.com/auth/calendar',
        'https://www.googleapis.com/auth/calendar.readonly',
    ],
    message=MISSING_CLIENT_SECRETS_MESSAGE)


class MainPage(webapp2.RequestHandler):
    def get(self):
        template_values = {
            'imported_event_count': self.request.get('imported_event_count', -1)
        }
        template = JINJA_ENVIRONMENT.get_template('index.html')
        self.response.write(template.render(template_values))


class Importer(webapp2.RequestHandler):
    @decorator.oauth_required
    def get(self):
        template_values = {
            'imported_event_count': self.request.get('imported_event_count', -1)
        }
        template = JINJA_ENVIRONMENT.get_template('import.html')
        self.response.write(template.render(template_values))

    @decorator.oauth_required
    def post(self):
        http = decorator.http()
        calendar_id = self.request.get('calendar_id')
        csv_string = self.request.get('csv_file')
        jsonEvents = CsvEventLoader(StringIO.StringIO(csv_string)).jsonEvents()
        for jsonEvent in jsonEvents:
            try:
                service.events().insert(calendarId=calendar_id, body=jsonEvent).execute(http=http)
            except client.AccessTokenRefreshError:
                self.response.write(
                    """<!doctype html><html><body>
                    The credentials have been revoked or expired, please re-run
                    the application to re-authorize
                    </body></html>""")

        self.redirect('/?' + urllib.urlencode({'imported_event_count': len(jsonEvents)}))


class CsvEventLoader:
    def __init__(self, csv_file):
        self.event_list = csv.DictReader(csv_file)

    def jsonEvents(self):
        return [self._toJsonEvent(event) for event in self.event_list]

    def _toJsonEvent(self, event):
        start_dateTime = event['Start Date'] + 'T' + event['Start Time'] + 'Z'
        end_dateTime = event['End Date'] + 'T' + event['End Time'] + 'Z'
        return {
            'summary': event['Summary'],
            'description': event['Description'],
            'location': event['Location'],
            'start': {
                'dateTime': start_dateTime
            },
            'end': {
                'dateTime': end_dateTime
            },
            'attendees': self._toJsonAttendees(event['Attendees'])
        }

    def _toJsonAttendees(self, attendees):
        return [
            {'email': attendee} for attendee in str.split(attendees)
        ]


application = webapp2.WSGIApplication([
                                          ('/', MainPage),
                                          ('/import', Importer),
                                          (decorator.callback_path, decorator.callback_handler())
                                      ], debug=True)
