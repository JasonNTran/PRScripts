import argparse
from dataclasses import dataclass
import itertools
import os
import math
import re
from typing import Optional
import openpyxl
import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageChops
from collections.abc import Generator

# Settings
BG_PATH = './Template/genshin_bg.png'
FRAME_PATH = './Template/seasons_frame.png'
AVATAR_PATH = './avatars'
HONORABLE_PATH = './honorables'
SAVE_PATH = "./PR"
AVATAR_SIZE = 120
ALBUM_ART_SIZE = 120
FRAME_RIGHT_PADDING = 20
# Where each box in the frame starts and ends based on pixel percentage. Hardcoded for this template. Tuples are (left, top), (right, bottom)
YEAR_POSITION = [(0.096793003, 0.003102378), (0.29271137, 0.065149948)]
SEASON_POSITION = [(0.402332362, 0.003102378), (0.598250729, 0.065149948)]
TYPE_POSITION = [(0.70728863, 0.003102378), (0.903206997, 0.065149948)]
SCORE_POSITION = [(0.401749271, 0.932781799), (0.597667638, 0.994829369)]
""" The sheet several columns to work properly. It needs:
- An anime column containing "Anime" (Ex. Anime, Anime Info, Anime Name)
- A song name column called one of the following: 'song info', 'songinfo', 'songartist', "song name", "songname"
- A total column "Total"
- A rank column "Rank"
- A nominator column "Nominator"
- A column for each ranker with their name matching their avatar icon file name
"""
DOC_STRING = """
Usage:
  generate_panels.py <sheet path.xlsx> <people_inside_video_box (OPTIONAL, INT)>
"""


class FontStyles:
    @staticmethod
    def load_fonts():
        return {
            "song": ImageFont.truetype("Fonts/Montserrat-Semibold.ttf", size=36),
            "score": ImageFont.truetype("Fonts/Montserrat-Semibold.ttf", size=36),
            "anime": ImageFont.truetype("Fonts/Montserrat-Semibold.ttf", size=34),
            "type": ImageFont.truetype("Fonts/Montserrat-Semibold.ttf", size=30),
            "season": ImageFont.truetype("Fonts/Montserrat-Semibold.ttf", size=30),
            "year": ImageFont.truetype("Fonts/Montserrat-Semibold.ttf", size=30),
            "remaining_hms": ImageFont.truetype("Fonts/Montserrat-Semibold.ttf", size=24.03),
            "tokens": ImageFont.truetype("Fonts/Montserrat-Semibold.ttf", size=48.06),
            "name": ImageFont.truetype("Fonts/Montserrat-Semibold.ttf", size=24),
            "count_info": ImageFont.truetype("Fonts/Montserrat-Semibold.ttf", size=24.03)
        }


class PanelInfo:
    def __init__(self, info_dict: dict, save_path: str, name):
        self.index_dict = info_dict
        self.base_path = save_path
        self.fonts = FontStyles.load_fonts()

        self.background = Image.open(BG_PATH)
        self.video_frame = Image.open(FRAME_PATH)
        self.name = "Potato"

        self.positions = {
            "hm_count": None,
            "op_count": None,
            "ed_count": None,
            "in_count": None,
            "female_count": None,
            "male_count": None,
            "both_count": None,
        }

    @property
    def offset(self) -> tuple[int, int]:
        bg_w, bg_h = self.background.size
        vf_w, vf_h = self.video_frame.size
        offset_h = (bg_h - vf_h) // 2
        offset_w = bg_w - vf_w - FRAME_RIGHT_PADDING
        return (offset_w, offset_h)

    @property
    def frame_pos(self) -> list[int, int, int, int]:
        offset = self.offset
        vf_w, vf_h = self.video_frame.size
        return [offset[0], offset[0] + vf_w, offset[1], offset[1] + vf_h]

    def set_pos(self, **kwargs):
        for key, value in kwargs.items():
            if key in self.positions:
                self.positions[key] = value
            else:
                raise KeyError(f"Invalid position key: {key}")


def create_template(template_details: PanelInfo) -> Image.Image:
    frame_pos = template_details.frame_pos
    bg_w, bg_h = template_details.background.size
    template_image = Image.new('RGBA', (bg_w, bg_h))
    template_image.paste(template_details.background)
    template_image.paste(template_details.video_frame, template_details.offset)

    # Avatar and name
    draw = ImageDraw.Draw(template_image)
    start_x = (frame_pos[0] - AVATAR_SIZE) // 2
    start_y = frame_pos[2]
    path = os.path.join(AVATAR_PATH, f"{template_details.name}.png")
    avatar = Image.open(path).resize((AVATAR_SIZE, AVATAR_SIZE))
    template_image.paste(avatar, (start_x, start_y))

    draw.text(
        (start_x + AVATAR_SIZE / 2, start_y), template_details.name,
        font=template_details.fonts["name"], fill="white",
        stroke_width=2, stroke_fill='black',
        anchor="mm"
    )
    draw.rectangle(
        ((start_x, start_y), (start_x + AVATAR_SIZE, start_y + AVATAR_SIZE)),
        outline=(193, 193, 193), width=1
    )

    # Base for HM

    bbox = draw.multiline_textbbox(
        (0, 0),
        "Remaining\nHonorable\nMentions",
        font=template_details.fonts["count_info"],
        spacing=5)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]

    draw.multiline_text(((frame_pos[0] - text_w) // 2, bg_h // 2 - text_h),
                        "Remaining\nHonorable\nMentions",
                        font=template_details.fonts["count_info"],
                        fill="white", spacing=5, align="center")

    ed_bbox = draw.textbbox((0, int(bg_h * 0.6)),
                            "ED",
                            font=template_details.fonts["count_info"])
    op_bbox = draw.textbbox((0, int(bg_h * 0.6)),
                            "OP",
                            font=template_details.fonts["count_info"],)
    in_bbox = draw.textbbox((0, int(bg_h * 0.6)),
                            "IN",
                            font=template_details.fonts["count_info"])

    # Base for song type counts
    ed_length = ed_bbox[2] - ed_bbox[0]
    op_length = op_bbox[2] - op_bbox[0]
    in_length = in_bbox[2] - in_bbox[0]

    length_type_text = ed_length + op_length + in_length + 30
    op_x = (frame_pos[0] - length_type_text) // 2
    ed_x = op_x + op_length + 15
    in_x = ed_x + ed_length + 15
    line_start_x = op_x - 5
    line_end_x = op_x + length_type_text + 5
    draw.text((op_x, int(bg_h * 0.6)),
              "OP",
              font=template_details.fonts["count_info"],
              fill="white", align="center")

    draw.text((ed_x, int(bg_h * 0.6)),
              "ED",
              font=template_details.fonts["count_info"],
              fill="white", align="center")
    draw.text((in_x, int(bg_h * 0.6)),
              "IN",
              font=template_details.fonts["count_info"],
              fill="white", align="center")
    draw.rectangle((line_start_x, op_bbox[3] + 15, op_x + length_type_text + 5, op_bbox[3] + 18),
                   fill="white",
                   outline=(193, 193, 193), width=1)

    # Base for artist type
    artist_x = line_start_x + 5
    male_bbox = draw.textbbox(
        (0, 0), "Male", font=template_details.fonts["count_info"])
    female_bbox = draw.textbbox(
        (0, 0), "Female", font=template_details.fonts["count_info"])

    male_text_height = male_bbox[3] - male_bbox[1]
    female_text_height = female_bbox[3] - female_bbox[1]

    male_y_pos = int(bg_h * 0.75)
    male_line_y_pos = male_y_pos + male_text_height + 20
    draw.text((artist_x, male_y_pos), "Male",
              font=template_details.fonts["count_info"], fill="white")
    draw.rectangle((line_start_x, male_line_y_pos, line_end_x, male_line_y_pos + 3),
                   fill="white", outline=(193, 193, 193), width=1)

    female_y_pos = male_line_y_pos + 11
    female_line_y_pos = female_y_pos + female_text_height + 20
    draw.text((artist_x, female_y_pos), "Female",
              font=template_details.fonts["count_info"], fill="white")
    draw.rectangle((line_start_x, female_line_y_pos, line_end_x, female_line_y_pos + 3),
                   fill="white", outline=(193, 193, 193), width=1)

    both_y_pos = female_line_y_pos + 11
    draw.text((artist_x, both_y_pos), "Both",
              font=template_details.fonts["count_info"], fill="white")

    template_details.set_pos(
        hm_count=(frame_pos[0] // 2, bg_h // 2 + 15),
        op_count=(op_x + op_length // 2, op_bbox[3] + 30),
        ed_count=(ed_x + ed_length // 2, op_bbox[3] + 30),
        in_count=(in_x + in_length // 2, op_bbox[3] + 30),
        male_count=(line_end_x - 15,  male_y_pos),
        female_count=(line_end_x - 15, female_y_pos),
        both_count=(line_end_x - 15, both_y_pos))

    return template_image


def create_song_panel(
    row: tuple,
    template_details: PanelInfo,
    template_panel: Image.Image,
) -> None:
    index_dict = template_details.index_dict
    panel = template_panel.copy()
    song_info = {key: row[index_dict[key]] for key in list(index_dict.keys())}
    season = song_info['season']
    write_song_info(song_info, panel, template_details)
    write_count_info(song_info, panel, template_details)
    write_honorables(song_info, panel, template_details)
    season_dict = {
        "Winter": "1",
        "Spring": "2",
        "Summer": "3",
        "Fall": "4"
    }
    save_path = os.path.join(template_details.base_path,
                             f"panel_{int(song_info['year'])}_{season_dict[season]}_{season}.png")
    panel.save(save_path)


def write_song_info(
    song_info: dict,
    panel: Image.Image,
    template_details: PanelInfo
) -> None:
    offset = template_details.offset
    bg_w, bg_h = template_details.background.size
    vf_w, vf_h = template_details.video_frame.size

    song_name = str(song_info["song_info"])
    anime_name = str(song_info["anime"])
    year = str(int(song_info["year"]))
    season = str(song_info["season"])
    song_type = str(song_info["song_type"])
    score = song_info["score"]

    draw = ImageDraw.Draw(panel)
    song_font = template_details.fonts["song"]
    anime_font = template_details.fonts["anime"]

    draw.rectangle(
        ((offset), (offset[0] + vf_w, offset[1] + vf_h)), outline=(193, 193, 193), width=1)
    draw.text((offset[0] + vf_w / 2, offset[1] / 2), anime_name, font=song_font,
              fill='white', stroke_width=1, stroke_fill='black', anchor="mm")
    draw.text((offset[0] + vf_w / 2, (bg_h + offset[1] + vf_h) / 2), song_name, font=anime_font,
              fill='white', stroke_width=1, stroke_fill='black', anchor="mm")

    year_box_center = (offset[0] + (YEAR_POSITION[0][0] * vf_w + YEAR_POSITION[1][0] * vf_w) / 2,
                       offset[1] + (YEAR_POSITION[0][1] * vf_h + YEAR_POSITION[1][1] * vf_h) / 2)
    season_box_center = (offset[0] + (SEASON_POSITION[0][0] * vf_w + SEASON_POSITION[1][0] * vf_w) / 2,
                         offset[1] + (SEASON_POSITION[0][1] * vf_h + SEASON_POSITION[1][1] * vf_h) / 2)
    type_box_center = (offset[0] + (TYPE_POSITION[0][0] * vf_w + TYPE_POSITION[1][0] * vf_w) / 2,
                       offset[1] + (TYPE_POSITION[0][1] * vf_h + TYPE_POSITION[1][1] * vf_h) / 2)
    score_box_center = (offset[0] + (SCORE_POSITION[0][0] * vf_w + SCORE_POSITION[1][0] * vf_w) / 2,
                        offset[1] + (SCORE_POSITION[0][1] * vf_h + SCORE_POSITION[1][1] * vf_h) / 2)

    draw.text(year_box_center, year,
              font=template_details.fonts["year"], fill=(59, 60, 67),
              stroke_width=0, stroke_fill='black', anchor="mm")
    draw.text(season_box_center, season,
              font=template_details.fonts["season"], fill=(59, 60, 67),
              stroke_width=0, stroke_fill='black', anchor="mm")
    draw.text(type_box_center, song_type,
              font=template_details.fonts["type"], fill=(59, 60, 67),
              stroke_width=0, stroke_fill='black', anchor="mm")
    draw.text(score_box_center, score,
              font=template_details.fonts["score"], fill=(59, 60, 67),
              stroke_width=0, stroke_fill='black', anchor="mm")


def write_count_info(
    song_info: dict,
    panel: Image.Image,
    template_details: PanelInfo
):
    positions = template_details.positions
    draw = ImageDraw.Draw(panel)
    count_font = template_details.fonts["count_info"]
    hm_font = template_details.fonts["tokens"]
    draw.text(positions["hm_count"], str(int(song_info["tokens"])),
              font=hm_font, fill="white", anchor='ma',
              stroke_width=0, stroke_fill='black')
    draw.text(positions["op_count"], str(int(song_info["op_count"])),
              font=count_font, fill="white", anchor='ma',
              stroke_width=0, stroke_fill='black')
    draw.text(positions["ed_count"], str(int(song_info["ed_count"])),
              font=count_font, fill="white", anchor='ma',
              stroke_width=0, stroke_fill='black')
    draw.text(positions["in_count"], str(int(song_info["in_count"])),
              font=count_font, fill="white", anchor='ma',
              stroke_width=0, stroke_fill='black')
    draw.text(positions["male_count"], str(int(song_info["male_count"])),
              font=count_font, fill="white", anchor='ma',
              stroke_width=0, stroke_fill='black')
    draw.text(positions["female_count"], str(int(song_info["female_count"])),
              font=count_font, fill="white", anchor='ma',
              stroke_width=0, stroke_fill='black')
    draw.text(positions["both_count"], str(int(song_info["both_count"])),
              font=count_font, fill="white", anchor='ma',
              stroke_width=0, stroke_fill='black')


def write_honorables(
    song_info: dict,
    panel: Image.Image,
    template_details: PanelInfo
):
    honorable = song_info["honorables"]
    if honorable is not None and honorable != "":
        split = str(honorable).split("by")
        # try:
        #     path = os.path.join(HONORABLE_PATH, f"{song_name}.png")
        #     album_art = Image.open(path).resize(
        #         (AVATAR_SIZE, AVATAR_SIZE))
        # except FileNotFoundError:
        #     print(f"[WARNING] Could not find image for {song_name}")
        album_art = Image.new(
            'RGBA', (ALBUM_ART_SIZE, ALBUM_ART_SIZE), (0, 0, 0, 255))
        frame_pos = template_details.frame_pos
        start_x = frame_pos[0] + 10
        start_y = frame_pos[3] - AVATAR_SIZE - 30

        draw = ImageDraw.Draw(panel)
        panel.paste(album_art, (start_x, start_y))
        hm_bbox = draw.textbbox((0, 0),
                                "Honorable Mention",
                                font=template_details.fonts["song"])
        song_bbox = draw.textbbox((0, 0),
                                  str(honorable),
                                  font=template_details.fonts["anime"])
        hm_height = hm_bbox[3] - hm_bbox[1]
        total_height = hm_height + (song_bbox[3] - song_bbox[1])
        y_spacing = (ALBUM_ART_SIZE - total_height - 5) // 2
        draw.text((start_x + ALBUM_ART_SIZE + 15, start_y + y_spacing), "Honorable Mention:",
                  font=template_details.fonts["song"], fill="white",
                  stroke_width=2, stroke_fill='black')

        draw.text((start_x + ALBUM_ART_SIZE + 15, start_y + y_spacing + hm_height + 5), str(honorable),
                  font=template_details.fonts["anime"], fill="white",
                  stroke_width=2, stroke_fill='black')

    # last 30% of the line is the male/female count


def clean_name(name: str) -> str:
    name = re.sub(r'\d+', '', name)
    font = ImageFont.truetype("Fonts/antipasto.regular.ttf", size=30)
    name_length = font.getlength(name)
    while (name_length > AVATAR_SIZE):
        decrease = AVATAR_SIZE / name_length
        num_char = len(name)
        name = name[:int(num_char * decrease)]
        name_length = font.getlength(name)
    return name


def get_columns(sheet: openpyxl.worksheet.worksheet.Worksheet) -> dict:
    index_dict = {
        'year': None,
        "season": None,
        "anime": None,
        "song_link": None,
        "song_info": None,
        "song_type": None,
        "score": None,
        "op_count": None,
        "ed_count": None,
        "in_count": None,
        "tokens": None,
        "male_count": None,
        "female_count": None,
        "both_count": None,
        "honorables": None,
    }
    rules = [
        ("year", "year"),
        ("season", "season"),
        ("anime", "anime"),
        ("song link", "song_link"),
        ("type", "song_type"),
        ("score", "score"),
        ("op", "op_count"),
        ("ed", "ed_count"),
        ("in", "in_count"),
        ("tokens", "tokens"),
        ("male", "male_count"),
        ("female", "female_count"),
        ("both", "both_count"),
        ("honorary", "honorables"),
    ]

    for index, cell in enumerate(sheet[1]):
        column = cell.value
        if column:
            if column.lower() in ['song info', 'songinfo', 'songartist', "song name", "songname"]:
                if index_dict["song_info"] is None:
                    index_dict["song_info"] = index
                continue

            # normal cases
            for substr, key in rules:
                if index_dict[key] is None and substr.lower() in column.lower():
                    index_dict[key] = index
                    break
    return index_dict


def create_all_panels(sheet_name: str, save_path: str) -> None:
    wrkbk = openpyxl.load_workbook(sheet_name)
    sheet = wrkbk.active
    indices_info = get_columns(sheet)
    template_details = PanelInfo(indices_info, save_path, "Potato")
    template_panel = create_template(template_details)
    for index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True)):
        if row is None or row[0] is None:
            print(f"[INFO] Hit none on row {index + 2}. Exiting")
            break
        create_song_panel(row, template_details, template_panel)


def create_dirs(sheet_name: str) -> str:
    save_path = os.path.join(SAVE_PATH, "".join(
        sheet_name.split(".")[:-1]), "panels")
    if not os.path.exists(save_path):
        os.makedirs(save_path)
    return save_path


def main(sheet_name: str) -> None:
    save_path = create_dirs(sheet_name)
    file_path = os.path.join(os.getcwd(), sheet_name)
    create_all_panels(file_path, save_path)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Generate video panels from spreadsheet.')
    parser.add_argument('sheet', type=str, help='Path to the Excel sheet')
    args = parser.parse_args()
    main(args.sheet)
