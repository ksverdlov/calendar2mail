#!/usr/bin/env python
"""
Send a text report about events in Exchange calendar for the day.

Usage:
  calendar2mail.py
  calendar2mail.py 2020-08-01

References: https://github.com/juandesant/measure_calendar_occupancy
"""

import os
import sys
from exchangelib import Credentials, Configuration, Account, DELEGATE, Message, Mailbox, HTMLBody, EWSDateTime, EWSTimeZone, EWSDate
from datetime import timedelta
import yaml
# requires having used keyring.set_password(pwdkey, user, the_password)
import keyring

ewsfile = os.environ["HOME"]+"/.ewscfg.yaml"

# Read configuration data from ewscfg.yaml file
with open(ewsfile) as f:
    cfg_data = yaml.load(f, Loader=yaml.FullLoader)

email = cfg_data['EWS_EMAIL']  # Default email
user = cfg_data['EWS_USER']   # Default user
server = cfg_data['EWS_SERVER']  # Default server value
# Allows searching of the password in the keyring
pwdkey = cfg_data['EWS_PWDKEY']
emailto = cfg_data['EWS_EMAIL_TO']
subject = cfg_data['EWS_SUBJECT']
pwd = keyring.get_password(pwdkey, user)

creds = Credentials(username=user, password=pwd)
config = Configuration(server=server, credentials=creds)
account = Account(
    primary_smtp_address=email, config=config, autodiscover=False, access_type=DELEGATE
)

tz = EWSTimeZone('Europe/Moscow')

if len(sys.argv) > 1:
    report_date = EWSDateTime.strptime(sys.argv[1], '%Y-%m-%d').replace(tzinfo=tz)
else:
    report_date = EWSDateTime.today().replace(tzinfo=tz)

items = account.calendar.view(
    start = report_date + timedelta(hours=6),
    end = report_date + timedelta(hours=30)
)

duration = timedelta(hours=0)
list = ""
for item in items:
    duration += item.end - item.start
    info = []
    if item.location is not None:
        info.append(item.location)
    if item.required_attendees is not None:
        for person in item.required_attendees[:5]:
            info.append(str(person.mailbox.name))
        if len(item.required_attendees) > 5:
            info.append("…")
    if item.optional_attendees is not None:
        for person in item.optional_attendees[:2]:
            info.append(str(person.mailbox.name))
        if len(item.optional_attendees) > 2:
            info.append("…")
    list += "%s %s %s %s%s\n" % (item.start.astimezone(tz).strftime("%H:%M"), item.end.astimezone(tz).strftime("%H:%M"), '{0: <5}'.format(item.duration[2:]), item.subject, ' (' + ', '.join(info) + ')' if info else '')

body = "%s report\n\n" % report_date.strftime("%Y-%m-%d")
body += "Total time: %s\n\n" % duration
body += list

print(body)

confirm = input('Please confirm sending report to %s: [y/n]' % emailto)

if not confirm or confirm[0].lower() != 'y':
    print('You did not indicate approval')
    exit(1)

m = Message(
    account=account,
    folder=account.sent,
    subject=subject,
    body = HTMLBody('<html><body><pre>%s</pre></body></html>' % body),
    to_recipients = [Mailbox(email_address=emailto)]
)
m.send_and_save()