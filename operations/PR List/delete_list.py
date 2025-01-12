import requests
import time

class NotFoundException(Exception):
    pass

def read_query(filename):
    filepath = f"./queries/{filename}.graphql"
    f = open(filepath, 'r')
    return f.read()

def execute_query(query, variables, token=None):
    ENDPOINT = "https://graphql.anilist.co"
    headers = {}
    if token is not None:
        headers['Authorization'] = f"Bearer {token}"
    while True:
        r = requests.post(ENDPOINT, headers=headers, json={'query': query, 'variables': variables})
        if r.status_code == 200:
            res = r.json()
            if 'errors' in res:
                raise Exception(f"query response contained errors: {res['errors']}")         # using base exception for now
            return res['data']
        elif r.status_code == 429:
            timeout = float(r.headers.get('retry-after', 60))
            time.sleep(timeout / 1000.0)
        elif r.status_code == 404:
            raise NotFoundException
        else:
            raise Exception(f"query response returned {r.status_code}: {r.text}")

def clear_planning(token, username):
    get_media = read_query("get_planning")
    delete_media = read_query("delete_list_entry")
    while True:
        data = execute_query(get_media, variables={'username': username})['Page']
        num_entries = len(data['mediaList'])
        if num_entries == 0:
            break
        for entry in data['mediaList']:
            listId = entry['id']
            # print(f"Deleting entry: {{entry['media']['title']['romaji']}")
            data = execute_query(delete_media, variables={'mediaId': listId}, token=token)['DeleteMediaListEntry']
            deleted = data['deleted']
            if not deleted:
                print(f"Failed to delete entry: {listId} ({entry['media']['title']['romaji']})")
                exit()
        print(f"Deleted {num_entries} entries.")

def clear_all(token, username):
    get_media = read_query("get_media_list")
    delete_media = read_query("delete_list_entry")
    while True:
        data = execute_query(get_media, variables={'username': username})['Page']
        num_entries = len(data['mediaList'])
        if num_entries == 0:
            break
        for entry in data['mediaList']:
            listId = entry['id']
            # print(f"Deleting entry: {{entry['media']['title']['romaji']}")
            data = execute_query(delete_media, variables={'mediaId': listId}, token=token)['DeleteMediaListEntry']
            deleted = data['deleted']
            if not deleted:
                print(f"Failed to delete entry: {listId} ({entry['media']['title']['romaji']})")
                exit()
        print(f"Deleted {num_entries} entries.")

# username = "dreamclimber"
# token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsImp0aSI6IjhmOGRjMDNjNjgwZDJhNGE5NDg3NzJkZDdhNDc4YzA2Yjk5Y2M4ODFhZTFjZTQ0NDA3Zjc1MDAwZTA3ZmJjMWE0ZTYxOTZhZGQ1NDRkMTQ2In0.eyJhdWQiOiI2ODA3IiwianRpIjoiOGY4ZGMwM2M2ODBkMmE0YTk0ODc3MmRkN2E0NzhjMDZiOTljYzg4MWFlMWNlNDQ0MDdmNzUwMDBlMDdmYmMxYTRlNjE5NmFkZDU0NGQxNDYiLCJpYXQiOjE2ODk4OTUzMDksIm5iZiI6MTY4OTg5NTMwOSwiZXhwIjoxNzIxNTE3NzA5LCJzdWIiOiI1OTQxNzMyIiwic2NvcGVzIjpbXX0.lNoIF1qsnpyXRqjH6Ms97t6h9AgeUWDUv_ks6ygawVP34oBgAwN27MFJ1IEPIKG7UPsjgOayEeML1DxF15V-NX36-ljp5azCWCGPs_kc40cRv7YD66esWqWOchuXnTrj0-P8ISl48ZUT5GTQrg1u7m7mg_x-e8gmPIkKZnB3oyWfrF3TsCYAVkDKfTpEMJP3QlvPdR7kI1lkrbqvIyboqZu1T2F8g0t0DQ_B-JEzBlP_DqGLWN8fHZJ4H7Bio8p2KbA6ef8ILy5AiJRO6byQHvvsSMn76eZ30Oqb8wSsMZOz5EgNfJ_L7VYJnI6DMzU0yKjE9BCnLuhzLMAgbAhGOPfRCAWIubPoKo7YeDqxjfdFkOWI8EZM_My-JzltjBrdwmh_UNRZYOtqqYvm0WnPYtqcIWHypjOFl-1MTdHN_izZl7hMAn7Rjv6vTF6gNenl2STFJSQQi9gtMdFQoB4qiC20eONVs4jgtX0ImYYxY4pATKmfreFaOnANlpOOMErWr9lN1nNkTt7nFNTj8-rCi82WeKCy6GGP4quUZLE7a6ikP8MrWfnZOejocv2hRiAuw1GlVORtY3nvt1STrYCBMZcKEnsJCud5fr7TefnQ6lgdP70jnFJUvlfn2ML2DTu0QmC6GN3sdHQphIPybAxy7GbXvN8IHfOhmgkYdGX24uY"
username = "traininglist83"
token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsImp0aSI6IjI5OTEyYjMyNjBmNmM2ODA5ZmY5YTZkOGMzNzgzOTNiMTA4NGQyNjk2MzA3NDA3MzE4YWIxZDJlODI2OWMwZGNjMThhYmNkODljNmMyNDBlIn0.eyJhdWQiOiIxOTUxMyIsImp0aSI6IjI5OTEyYjMyNjBmNmM2ODA5ZmY5YTZkOGMzNzgzOTNiMTA4NGQyNjk2MzA3NDA3MzE4YWIxZDJlODI2OWMwZGNjMThhYmNkODljNmMyNDBlIiwiaWF0IjoxNzIxMjc5ODY4LCJuYmYiOjE3MjEyNzk4NjgsImV4cCI6MTc1MjgxNTg2OCwic3ViIjoiNjI0NzA2MSIsInNjb3BlcyI6W119.r8qhm3Wp6KITAd-ayAi-Bb938suU7sjrQgkcglLbzPHIQ6Iu10Y7e_7mcNwwFI4-HEpAC9p-rDySNxBDNzG2ldZ6bmnFVBp8x0Zm1XgcuTLqDxwi0sGm4QCyBEBCxl6B92xbl65-IvS2MEziuvoicKgaDHzrgNjquTWZ4jejZd-Hjb8-u8gSX8uUmCf3CzMp2FMMR0k6WzVlxrwyas-b5jTSd-pc9Sbb9hopz0rPiICLIqzyskpYZH0CTYWWc7AiOHUqWmjNXGaiv5_7-vhu6y-ubICsBfzcoz8yIMJPBHOp20PEFt-5vsXTVs8bGoeIpqYnmGztuLrQbBtuCw347-NTunZwXGN9QxbrVZ__SMF7vVt6bQup-wS3JAyqdCIHCFWLbTLoKD6-xgfYP6J9DJ4jTQYqwkcZFfr8tSNdHFrJwWqNvdPva5NAPYNmK5JahklFeQ4op-vgRU7nZUDzUnfXk1xVm5JGedgU1IBbWj21tpDaw34iItMd8e9C3TjKcfMAMvYMdkagb4i_K6BuTurs4Z3J4Dqris4lqGWXcOtqT0prDo__wHwNgLHCWRBDEEzgriYPf8M010kGjVowFSrnS6D3fwuuaX5-IwcJs-MlV4KGkYANrfSPhyI7Hoc4XVY4-G3CBxJBAmCeFDz6OfjJ6XLegzqP5GhdgIgJseA"
clear_all(token, username)
