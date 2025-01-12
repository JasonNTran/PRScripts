import requests
import time
from sqlalchemy import create_engine, sql, Table, String, Integer, ForeignKey, Column, MetaData

DOC_STRING = """
Usage:
 anilist_operations.py add-client <username> <client-id> <client-secret> <client-name> <redirect-url>
 anilist_operations.py add-token <username> <client-id>
 anilist_operations.py clear-list <username>
"""
MAX_RETIRES = 5

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
            print("[TIMEOUT] Starting calls to Anilist again")
        elif r.status_code == 404:
            return None
        else:
            raise Exception(f"[ERROR] Response returned {r.status_code}: {r.text}")

def get_token(username, conn):
    s = sql.select(user_table).where(user_table.c.username == username)
    row = conn.execute(s).fetchone()
    return row.token

def clear_list(token, username):
    get_all_media_query = '''
    query ($userName: String) { # Define which variables will be used in the query (id)
        MediaListCollection (userName: $userName, type: ANIME) {
            lists {
                entries {
                    id
                    media { title { romaji }}
                }
            }
        }
    }
    '''

    delete_media_query= '''
        mutation ($mediaId: Int) {
            DeleteMediaListEntry (id: $mediaId) {
                deleted
            }
        }
    '''
    media_lists = handleRequest(get_all_media_query, variables={'userName': username})['MediaListCollection']['lists']
    data  = []
    for list in media_lists:
        entry_list = list['entries']
        data.extend(entry_list)
    num_entries = len(entry_list)
    for entry in data:
        listId = entry['id']
        print(f"Deleting entry: {entry['media']['title']['romaji'].encode('utf-8')}")
        data = handleRequest(delete_media_query, variables={'mediaId': listId}, token=token)['DeleteMediaListEntry']
        deleted = data['deleted']
        if not deleted:
            print(f"Failed to delete entry: {listId} ({entry['media']['title']['romaji']})")
            exit()
    print(f"Deleted {num_entries} entries.")

def get_auth_url(client_id, client_table, conn):
    s = sql.select(client_table).where(client_table.c.id == client_id)
    row = conn.execute(s).fetchone()

    if row is None:
        return None
    user = row._asdict()
    redirect_uri = user['redirect_uri']
    uri = f"https://anilist.co/api/v2/oauth/authorize?client_id={client_id}&redirect_uri={redirect_uri}&response_type=code"
    user['auth_url'] = uri
    return user

def init_sql():
    SQL_DB_URL = "sqlite:///auth.db"
    engine = create_engine(SQL_DB_URL)
    metadata_obj = MetaData()
    client_table = Table("client",
                         metadata_obj,
                         Column("id", Integer, primary_key = True),
                         Column("secret", String),
                         Column("redirect_uri", String),
                         Column("app_name", String),
                         Column("username", String))
    user_table = Table("user",
                        metadata_obj,
                        Column("username", String, primary_key = True),
                        Column("client_id", Integer, ForeignKey('client.id')),
                        Column("token", String))
    metadata_obj.create_all(engine)
    return(engine, metadata_obj)

def convert_code_to_token(code, client):
    ENDPOINT =  'https://anilist.co/api/v2/oauth/token'

    params =  {
        'grant_type': 'authorization_code',
        'client_id': client['id'],
        'client_secret': client['secret'],
        'redirect_uri': client['redirect_uri'],
        'code': code, 
      }

    r = requests.post(ENDPOINT, json = params)
    if r.status_code == 200:
        response = r.json()
        print(response)
        return response['access_token']
    else:
        raise Exception(f"[ERROR] Response returned {r.status_code}: {r.text}")

if __name__ == '__main__':
    import sys
    try: 
        command = sys.argv[1]
    except IndexError:
        print(DOC_STRING)
        exit()
    if command == 'add-client':
        if (len(sys.argv) < 7):
            print(DOC_STRING)
            exit()
        username = sys.argv[2]
        client_id = sys.argv[3]
        client_secret = sys.argv[4]
        client_name = sys.argv[5]
        redirect_uri = sys.argv[6]

        engine, metadata = init_sql()
        with engine.begin() as conn:
            client_table = metadata.tables['client']
            s = sql.insert(client_table).values(
                id = client_id,
                secret = client_secret,
                redirect_uri = redirect_uri,
                app_name = client_name,
                username = username
            )
            conn.execute(s)
    
    elif command == 'add-token':
        if (len(sys.argv) < 4):
            print(DOC_STRING)
            exit()
        username = sys.argv[2]
        client_id = sys.argv[3]

        engine, metadata = init_sql()
        with engine.begin() as conn:
            client_table = metadata.tables['client']
            client = get_auth_url(client_id, client_table, conn)
            if client is None:
                print(f"client id {client_id} not found in table")
                s = sql.select(client_table)
                res = conn.execute(s)
                print("available clients: ")
                for row in res:
                    print(row)
                exit()
            code = input(f"authorize with this url on {username}: \n{client['auth_url']}\n")

            token = convert_code_to_token(code, client)
            user_table = metadata.tables['user']
            s = sql.insert(user_table).values(
                client_id = client_id,
                username = username,
                token = token
            )
            conn.execute(s)
    elif command == 'clear-list':
        if (len(sys.argv) < 3):
            print(DOC_STRING)
            exit()
        username = sys.argv[2]
        engine, metadata = init_sql()

        with engine.begin() as conn:
            user_table = metadata.tables['user']
            token = get_token(username, conn)
            clear_list(token, username)
