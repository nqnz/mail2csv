#!/usr/bin/env python3
import os
import json
import msal
import requests
from datetime import datetime, timezone

# ── Configuration ────────────────────────────────────────────────────────────
TENANT_ID      = os.environ['AZ_TENANT_ID']
CLIENT_ID      = os.environ['AZ_CLIENT_ID']
CLIENT_SECRET  = os.environ['AZ_CLIENT_SECRET']
USER_EMAIL     = os.environ['CSV_INGEST_EMAIL'] 
DATA_DIR       = os.environ['DATA_DIR'] 
STATE_FILE     = os.environ['STATE_FILE']
REPORT_SUBJ    = os.environ['R_SUBJECT']
# ──────────────────────────────────────────────────────────────────────────────


def load_state():
    if os.path.isfile(STATE_FILE):
        with open(STATE_FILE) as f:
            return json.load(f)
    # no state yet → start at “now”
    return {'deltaLink': None}


def save_state(state):
    with open(STATE_FILE, 'w') as f:
        json.dump(state, f)


def get_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=f'https://login.microsoftonline.com/{TENANT_ID}',
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=[ 'https://graph.microsoft.com/.default' ])
    return result['access_token']


def fetch_new_messages(token, delta_link):
    headers = {'Authorization': f'Bearer {token}'}
    if delta_link:
        # continue from last delta
        resp = requests.get(delta_link, headers=headers).json()
    else:
        # first time: get all unread messages matching subject
        filter_q = ("isRead eq false and subject eq "
                    f"'{REPORT_SUBJ}'")
        url = (f'https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/'
               f'mailFolders/Inbox/messages?'
               f'$filter={filter_q}&$select=id&$top=50')
        resp = requests.get(url, headers=headers).json()

    return resp.get('value', []), resp.get('@odata.deltaLink')


def process_message(token, msg_id):
    headers = {'Authorization': f'Bearer {token}'}
    # fetch attachments
    att_url = (f'https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/'
               f'messages/{msg_id}/attachments')
    resp = requests.get(att_url, headers=headers).json().get('value', [])
    for a in resp:
        if a.get('@odata.type','').endswith('fileAttachment'):
            name = a['name']
            # prepend timestamp to avoid collisions
            ts = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')
            out_path = os.path.join(DATA_DIR, f"{ts}_{name}")
            with open(out_path, 'wb') as f:
                f.write(bytes(a['contentBytes'], 'utf-8'))

    # mark message as read so we don’t fetch it again
    patch_url = (f'https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/'
                 f'messages/{msg_id}')
    requests.patch(patch_url,
                   headers={**headers, 'Content-Type':'application/json'},
                   json={'isRead': True})


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    state = load_state()
    token = get_token()

    msgs, new_delta = fetch_new_messages(token, state.get('deltaLink'))
    for m in msgs:
        process_message(token, m['id'])

    # save the new deltaLink to pick up only changes next time
    if new_delta:
        state['deltaLink'] = new_delta
        save_state(state)


if __name__ == '__main__':
    main()
