# **PR Scripts**

A collection of scripts that do various actions from a PR (power ranking) sheet

---

- ## <ins>Anilist Operations</ins>

Handles various operations on Anilist.co that a user may want to perform with their account, such as deleting their entire list or creating an auth token

---


- ## <ins>Catbox DL</ins>

Downloads all the audio or video links in a PR sheet. Can be used to download videos from a result sheet to edit the result video or audio files from a regular rank listen locally

---

- ## <ins>PR List</ins>

Adds shows to a list based on the shows on a sheet. Will also make notes on the list what sheets have been used and why each show was added (song and PR). Requires an Anilist token, which can be obtained and stored through Anilist Operations

---

- ## <ins>Panel Generator</ins>

Creates all the panels needed for a PR result video based on a sheet. Background can be changed freely as long as its a 1920x1080 image. The video frame, if changed, will need changes to the box locations to accurate write the total score, rank, and song type.

Various settings at the top can be changed to improve the look and feel based on the template used and number of members. This includes number of columns per half, spacing between avatars, and avatar size. If a different video frame is used, the position of the boxes in the settings will also need to be changed 

The sheet several columns to work properly. It needs:
- An anime column containing "Anime" (Ex. Anime, Anime Info, Anime Name)
- A song name column called one of the following: 'song info', 'songinfo', 'songartist', "song name", "songname"
- A total column "Total"
- A rank column "Rank"
- A nominator column "Nominator"
- A column for each ranker with their name matching their avatar icon file name
---
