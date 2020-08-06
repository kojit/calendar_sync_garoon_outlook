from pathlib import Path
import base64
import datetime as dt
import dateutil
import json
import requests
from O365 import Account

CONFIG_FILE = 'calendar_sync_garoon_outlook.json'
WEEKS = 2
MAX_EVENT_NUM = 100

def get_period():
    now = dt.datetime.now().astimezone()
    #print(now.tzinfo, now.isoformat())
    end = now + dt.timedelta(weeks=WEEKS)
    return now, end


def get_garoon_events(cfg, now, end):
    cybozu_credential = cfg['CYBOZU_USER_NAME'] + ':' + cfg['CYBOZU_USER_PASSWORD']
    basic_credential = cfg['BASIC_AUTH_USER'] + ':' + cfg['BASIC_AUTH_PASSWORD']
    url = cfg['BASE_URL'] + 'events'
    basic_credentials = base64.b64encode(basic_credential.encode('utf-8'))
    headers = {
        'content-type': 'application/json',
        'X-Cybozu-Authorization': base64.b64encode(cybozu_credential.encode('utf-8')),
        'Authorization': 'Basic ' + basic_credentials.decode('utf-8')
    }
    params = {
        'limit': MAX_EVENT_NUM,
        'rangeStart': now.isoformat(),
        'rangeEnd': end.isoformat()
    }

    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()
    res_json = response.json()
    #print(json.dumps(res_json, indent=2))

    outlook_origin_events = {}
    events = {}
    for event in res_json['events']:
        if 'repeatId' in event:
            gid = event['id'] + '_' + event['repeatId']
        else:
            gid = event['id']
        #print('Garoon {} {}'.format(gid, event['subject']))
        if event['subject'].startswith('OID:'):
            oidpair = event['subject'].split()[0]
            outlook_id = oidpair.split(':')[1]
            outlook_origin_events[outlook_id] = event
        else:
            events[gid] = event
    return events, outlook_origin_events


def get_outlook_events(cfg, now, end):
    credential = (cfg['AZURE_APP_APPLICATION_ID'], cfg['AZURE_APP_CLIENT_SECRET'])
    account = Account(credential)
    if not account.is_authenticated:
        account.authenticate(scopes=['basic', 'calendar_all'])
    schedule = account.schedule()
    calendar = schedule.get_default_calendar()

    q = calendar.new_query('start').greater_equal(now)
    q.chain('and').on_attribute('end').less_equal(end)
    events = calendar.get_events(limit=100, query=q, include_recurring=True)
    """
    # we can only get 25 events, so I will get every weeks
    now = dt.datetime.now().astimezone()
    events = []
    for i in range(WEEKS):
        end = now + dt.timedelta(weeks=1)
        q = calendar.new_query('start').greater_equal(now)
        q.chain('and').on_attribute('end').less_equal(end)
        now = end
        events = events + list(calendar.get_events(limit=100, query=q, include_recurring=True))
    """

    garoon_origin_events = {}
    outlook_events = {}
    for event in events:
        #print('Outlook ' + event.subject)
        if event.subject.startswith('GID:'):
            gidpair = event.subject.split()[0]
            garoon_id = gidpair.split(':')[1]
            garoon_origin_events[garoon_id] = event
            print('Outlook - Garoon Origin Event ' + event.subject)
        else:
            outlook_events[event.object_id] = event
            print('Outlook - Outlook Origin Event ' + event.subject)
    return calendar, garoon_origin_events, outlook_events


def update_outlook_event(cfg, oevent, gid, gevent):
    subject = 'GID:' + gid + ' - ' + gevent['subject']
    if subject != oevent.subject:
        oevent.subject = subject
        oevent.body = cfg['EVENT_URL'] + (gid.split('_')[0] if '_' in gid else gid)

    start = dateutil.parser.parse(gevent['start']['dateTime'])
    if start != oevent.start:
        oevent.start = start

    end = dateutil.parser.parse(gevent['end']['dateTime'])
    if end != oevent.end:
        oevent.end = end

    if 'facilities' in gevent and len(gevent['facilities']) > 0:
        location = gevent['facilities'][0]['name']
        if not oevent.location or location != oevent.location['displayName']:
            oevent.location = location

    is_all_day = True if gevent['isAllDay'] == 'true' else False
    if is_all_day != oevent.is_all_day:
        oevent.is_all_day = is_all_day

    if oevent.is_reminder_on != False:
        oevent.is_reminder_on = False

    oevent.save()   # O365 module only updates if there is any changes


def main(cfg):
    start, end = get_period()
    try:
        garoon_events, outlook_origin_events = get_garoon_events(cfg, start, end)
        outlook_calendar, garoon_origin_events, outlook_events = get_outlook_events(cfg, start, end)
    except Exception as e:
        print(e)
        return

    ### Garoon -> Outlook
    # remove garoon origin event on outlook if it no longer exists.
    for key in list(garoon_origin_events.keys()):
        if key not in garoon_events:
            print('remove event {}'.format(key))
            garoon_origin_events[key].delete()
            del garoon_origin_events[key]

    # add/update garoon events to outlook
    for key, value in garoon_events.items():
        if key in garoon_origin_events:
            update_outlook_event(cfg, garoon_origin_events[key], key, value)
        else:
            print('add event - {} {}'.format(key, value['subject']))
            oevent = outlook_calendar.new_event()  # creates a new unsaved event
            update_outlook_event(cfg, oevent, key, value)

    ### TODO: Outlook -> Garoon
    """
    # remove outlook origin event on garoon if it no longer exists.
    for key in list(outlook_origin_events.keys()):
        if key not in outlook_events:
            print('remove event {}'.format(key))
            outlook_origin_events[key].delete()
            del outlook_origin_events[key]

    # add/update outlook events to garoon
    for key, value in outlook_events.items():
        if key in outlook_origin_events:
            update_garoon_event(cfg, outlook_origin_events[key], key, value)
        else:
            print('add event - {}'.format(value.subject))
            gevent = garoon_calendar.new_event()  # creates a new unsaved event
            update_garoon_event(cfg, gevent, key, value)
    """


if __name__ == '__main__':
    if Path.exists(Path.cwd() / CONFIG_FILE):
        with (Path.cwd() / CONFIG_FILE).open() as f:
            main(json.load(f))
    elif Path.exists(Path.home() / CONFIG_FILE):
        with (Path.home() / CONFIG_FILE).open() as f:
            main(json.load(f))
    else:
        print('There is no config file')
