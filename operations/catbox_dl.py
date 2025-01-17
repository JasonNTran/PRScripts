import os
import subprocess
import sqlite3
import requests
import openpyxl
from urllib.parse import urlparse

import sys

DOC_STRING = """
Usage: 
  catbox_dl.py <command> <sheet.xlsx> <exclude-artist (optional)>
"""
exclude_artists = 0

def cleanup_song(song): 
    if exclude_artists:
        song = "".join(song.split("by")[:-1])
        if ' by' in song:
            song = "".join(song.split("by")[:-1])
        elif ' BY' in song:
            song = "".join(song.split("BY")[:-1])
        else:
            print(f'[WARN] Could not split artist from song for song /{song.encode("utf-8")}/')
    song = song.replace("\"", "")
    song = song.replace("'", "")
    song = song.replace("*", "")
    song = song.replace("?", "")
    song = song.replace(":", "")
    song = song.replace("<", "-")
    song = song.replace(">", "-")
    song = song.strip()
    return song

def get_columns(sheet, isMp3): 
    song_column = None
    link_column = None
    rank_column = None
    for index, cell in enumerate(sheet[1]):
        column = cell.value
        if column:
            if column.lower() in ['song info', 'songinfo', 'songartist'] and song_column is None:
                song_column = index
            elif 'mp3' in column.lower() and isMp3:
                link_column = index
            elif column.lower() in ['songlink', 'song link'] and not isMp3:
                link_column = index
            elif "rank" in column.lower():
                rank_column = index
    if link_column == None:
        link_column = song_column
    
    return (song_column, link_column, rank_column)

def dl_song(hostname, link, file_name, isMp3):
    if hostname in ["www.youtube.com", "youtu.be"]:
        out_path = f"{file_name}.%(ext)s"
        subprocess.run([
            'yt-dlp',
            '-x', '-q',
            '--encoding', 'utf-8',
            '-o', f'{out_path}',
            link
            ], encoding = 'utf-8')
    elif hostname in ["files.catbox.moe", "openings.moe", "ladist1.catbox.video", "naedist.animemusicquiz.com"]:
        response = requests.get(link)
        extension = link.split(".")[-1]
        fileName = f"{file_name}.{extension}"
        with open(fileName, "wb") as file:
            file.write(response.content)
        if isMp3 and extension == "webm":
            subprocess.run(['ffmpeg', '-i', fileName, f"{file_name}.mp3"])
            os.remove(fileName)
    else:
        print("Hostname not recognized", hostname, file_name)
        sys.stdout.flush() 
    
def dl_pr(file_name):
    folder = os.path.splitext(file_name)[0]
    if not os.path.exists(folder):
        os.mkdir(folder)
        
    wrkbk  = openpyxl.load_workbook(file_name)
    sheet = wrkbk.worksheets[0]
    song_column, link_column, rank_column = get_columns(sheet, True)
    
    for row in sheet.iter_rows(min_row = 2):
        song_name = row[song_column].value
        rank = (int) (row[rank_column].value)
        try:
            print("method 1")
            link = row[link_column].hyperlink.target
        except:
            try:
                link = row[song_column].hyperlink.target
            except:
                link, song_name = song_name.split('by')[1::2]
        song_name = cleanup_song(song_name)
        host_name = urlparse(link).hostname
        file_name = f"{folder}/{rank}-{song_name}"
        dl_song(host_name, link, file_name, True)

def dl_result(file_name):
    folder = os.path.splitext(file_name)[0]
    if not os.path.exists(folder):
        os.mkdir(folder)
        
    wrkbk  = openpyxl.load_workbook(file_name)
    sheet = wrkbk.worksheets[0]
    song_column, link_column,_ = get_columns(sheet, False)
    count = 2
    for row in sheet.iter_rows(min_row = 2):
        song_name = row[song_column].value
        if song_name is None:
            print(f"[INFO] Song name is none on row {count}. Exiting loop")
            break
        try:
            link = row[link_column].hyperlink.target
        except:
            link, song_name = song_name.split('by')[1::2]
        song_name = cleanup_song(song_name)
        host_name = urlparse(link).hostname
        file_name = f"{folder}/{song_name}"
        dl_song(host_name, link, file_name, False)
        count += 1


if __name__ == '__main__':
    try:
        command = sys.argv[1]
        file = sys.argv[2]
    except IndexError:
        print(DOC_STRING)
        exit()
    if len(sys.argv) == 3:
        exclude_artists = sys.argv[2]
    if command == 'dl_result':
        dl_result(file)
    elif command == 'dl_pr':
        dl_pr(file)
    else:
        print(DOC_STRING)
        exit()

    