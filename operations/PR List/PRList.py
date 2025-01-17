# -*- coding: utf-8 -*-
import sys
import json
import os
import time
import pandas as pd
import requests
from sqlalchemy import create_engine, sql, Table, MetaData

DOC_STRING = """
Usage: 
  PRLIST.py <username> <sheet directory>
"""
ANILIST_ENDPOINT = 'https://graphql.anilist.co'
MAX_RETIRES = 5

def cleanup_song(song):
    split_string = ""
    if ' by' in song:
        split_string = ' by'
    elif ' BY' in song:
        split_string = ' BY'
    else:
        raise Exception(f'[ERROR] Could not split artist from song for song /{song}/')
    
    song = "".join(song.split(split_string)[:-1])
    if song[0] == "\"":
        song = song[1:]
    if song[-1] == "\"":
        song = song[:-1]
    return song

def handleRequest(query, variables, token=None):
    ENDPOINT = "https://graphql.anilist.co"
    count = 0
    headers = {}
    if token is not None:
        headers['Authorization'] = f"Bearer {token}"
    params = {
        'query': query,
        'variables': variables
    }
    while True:
        r =  requests.post(ENDPOINT, headers = headers, json = params)
        if r.status_code == 200:
            response = r.json()
            if 'errors' in response:
                print(f"[WARNING] query response contained errors: {response['errors']}")
            return response['data']
        elif r.status_code == 429:
            if count == MAX_RETIRES:
                print(f"[ERROR] Reached max timeouts")
                exit()
            print(f"[TIMEOUT] 429 Too Many Requests, sleeping", flush=True)
            timeout = float(r.headers.get('Retry-After', 60))
            count = count + 1
            time.sleep(timeout)
            print("[TIMEOUT] Starting calls to Anilist again", flush=True)
        elif r.status_code == 404:
            return None
        else:
            raise Exception(f"[ERROR] Response returned {r.status_code}: {r.text}")
        
def get_user_about(username):
    query = '''
        query ($userName: String) {
            User (name: $userName) {
                about
            }
        }
    '''
    variables = {
        'userName': username
    }
    response = handleRequest(query, variables)
    if response['User']['about'] is None:
        return []
    return response['User']['about'].split("\n")

def write_user_about(username, user_about, sheets_list, token):
    string_about = "\n".join(user_about) + "\n"
    for sheet in sheets_list:
        string_about += f"{sheet.replace(".xlsx", "")}\n"
    query = ''' 
        mutation ($about: String) {
            UpdateUser (about: $about) {
                about
            }
        }
    '''

    variables = {
        'userName': username,
        'about': string_about
    }
    
    print(f"[INFO] Writing to User's about")
    handleRequest(query, variables, token)

def add_anime(anilist_id, anime_info, media_entry, notes, token):
    notes = handleNotes(media_entry, anime_info)
    variables = {
        'notes': notes
    }
    if media_entry is not None:
        variables['entryId'] = media_entry['id']
    else:
        variables['mediaId'] = anilist_id

    query = '''
        mutation ($entryId: Int, $mediaId: Int, $notes: String) {
            SaveMediaListEntry (id: $entryId, mediaId: $mediaId, notes: $notes, status: COMPLETED) {
                id
                notes
            }
        }
    '''
    print(f"[INFO] Adding show: {anime_info['anime_name'].encode('utf-8')}", flush=True)
    r = handleRequest(query, variables, token)
    
def get_all_entries(user_name):
    query = '''
        query ($userName: String) { # Define which variables will be used in the query (id)
            MediaListCollection (userName: $userName, status: COMPLETED, type: ANIME) {
                lists {
                    entries {
                        id
                        notes
                        mediaId
                    }
                }
            }
        }
    '''
    variables = {
        'userName': user_name,
    }

    response = handleRequest(query, variables)
    if response is not None:
        return response['MediaListCollection']['lists']
    return None
    
def handleNotes(media_entry, anime_info):
    note_to_add = ""
    current_notes = ""
    if media_entry is not None and media_entry['notes'] is not None:
        current_notes =  media_entry['notes'] +"\n"
    for pr in anime_info['prs']:
        note_to_add += f"{pr}\n"
        note_to_add += "-" * 46 + "\n"
        for song in anime_info['prs'][pr]:
            note_to_add += f"- {song}\n"
        note_to_add += "\n"
    if note_to_add not in current_notes:
        current_notes += note_to_add
    return current_notes
    
def get_anilist_id(show, song, sheet):
    ENDPOINT = "https://anisongdb.com/api/search_request"

    params = {
        "anime_search_filter" : { "search" : show, "partial_match" : False },
        "song_name_search_filter" : { "search" : song, "partial_match" : False },
        "and_logic" : True,
    }
    r = requests.post(ENDPOINT, json=params)
    if r.status_code == 200:
        parsed = json.loads(r.content)
        if not parsed:
            return None
        for item in parsed:
            print(item.encode)
        return parsed[0]["linked_ids"]["anilist"]
    else:
        raise Exception(f"[ERROR] Response returned {r.status_code}: {r.text}")

def get_all_sheets(file_path, user_about):
    sheets_list = []
    for file in os.listdir(file_path):
        if file.endswith('.xlsx'):
            sheets_list.append(file)
    # for sheet in user_about:
    #     sheet_name = sheet + ".xlsx"
    #     if sheet_name in sheets_list:
    #         sheets_list.remove(sheet_name)
    return sheets_list

def get_anime_from_sheet(file_path, sheets_list):
    anime_list = {}
    manual_add = []
    for sheet_name in sheets_list:
        print(f"[INFO] Adding shows for: {sheet_name.encode('utf-8')}", flush=True)
        sheet_path = os.path.join(file_path, sheet_name)
        df = pd.read_excel(sheet_path)
        anime_column = None
        song_column = None
        for column in df.columns.tolist():
            if "anime" in column.lower() and anime_column is None:
                anime_column = column
            elif ('song info' in column.lower() or 'songinfo' in  column.lower() or 'songartist' in  column.lower()) and song_column is None:
                song_column = column
        start = time.time()
        for index, row in df.iterrows():
            pr = sheet_name[:-5]
            anime_name = row[anime_column]
            song_name = cleanup_song(row[song_column])
            anilist_id = get_anilist_id(anime_name, song_name, pr)
            if anilist_id is None:
                entry = {
                    'anime_name': anime_name,
                    'song_name': song_name,
                    'pr': pr
                }
                manual_add.append(entry)
            else:
                entry = anime_list.get(anilist_id)
                if entry is None:
                    entry = {
                        'anime_name': anime_name,
                        'prs': {}
                    }
                    anime_list[anilist_id] = entry
                pr_song_list = entry['prs'].get(pr)
                if pr_song_list is None:
                    entry['prs'][pr] = []
                entry['prs'][pr].append(song_name)

        end = time.time()
        print(f"[INFO] Looked up {len(df)} shows in {end - start}", flush=True)
    return (anime_list, manual_add)
      
def generate_anilist(file_path, user_name, token):
    user_about = get_user_about(user_name)
    sheets_list = get_all_sheets(file_path, user_about)
    anime_list, manual_add = get_anime_from_sheet(file_path, sheets_list)
    if anime_list is not None and len(sheets_list) != 0:
        entries = get_all_entries(user_name)
        entry_dict = {}
        if entries:
            entry_list = entries[0]['entries']
            entry_dict = dict((i['mediaId'], {'id' : i['id'], 'notes': i['notes']}) for i in entry_list)

        for anilist_id in anime_list:
            anime_info = anime_list[anilist_id]
            media_entry = entry_dict.get(anilist_id, None)
            notes = handleNotes(media_entry, anime_info)
            add_anime(anilist_id, anime_info, media_entry, notes, token)

        write_user_about(user_name, user_about, sheets_list, token)
        print("[INFO] Finished adding shows.")
        if len(manual_add) != 0:
            print("Listing shows to add manually")
            file = open("manual_add.txt", "a", encoding='utf-8')
            for entry in manual_add:
                line = (f"[INFO] Add show / {entry['anime_name'].encode('utf-8')} /" +
                        f"for song /{entry['song_name'].encode('utf-8')} /" + 
                        f"in pr / {entry['pr'].encode('utf-8')} /\n")
                file.write(line)
                print(line)
    else:
        print("[INFO] Nothing to add. Exiting")

def get_token(user_name):
    SQL_DB_URL = "sqlite:///auth.db"
    engine = create_engine(SQL_DB_URL)
    conn = engine.connect()

    with engine.begin() as conn:
        user_table = Table('user', MetaData(), autoload_with=engine)
        s = sql.select(user_table).where(user_table.c.username == user_name)
        result = conn.execute(s)
        user = result.fetchone()
        return user.token
 
if __name__ == '__main__':
    import sys
    try:
        user_name  = sys.argv[1]
        file_path = os.path.join(os.getcwd(), sys.argv[2])
    except IndexError:
        print(DOC_STRING)
        exit()
    token = get_token(user_name)
    generate_anilist(file_path, user_name, token)
