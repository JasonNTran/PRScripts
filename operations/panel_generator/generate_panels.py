import itertools
import os
import math
import re
from typing import Optional
import openpyxl
import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageChops

from collections.abc import Generator

#Settings
BG_PATH = './Template/black.png'
FRAME_PATH = './Template/frame_cropped.png'
AVATAR_PATH = './avatars'
SAVE_PATH = "./PR"
X_PAD = 25
Y_PAD = 30
MIN_EDGE_PADDING = 40
AVATAR_SIZE = 105
NUM_COLS_PER_SIDE = 2
# Where each box in the frame starts and ends based on pixel percentage. Hardcoded for this template
RANK_POSITION = [(0.096969697, 0.023784902), (0.172727273, 0.119958635)] 
TOTAL_POSITION = [(0.853787879, 0.883143744), (0.928787879, 0.979317477)] 
TYPE_POSITION = [(0.785606061, 0.006204757), (0.995454545, 0.055842813)] 

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
  generate_panels.py <sheet path.xlsx>
"""
class Person: 
    def __init__(self, name: str, index: int, image: Image.Image):
        self.full_name = name
        self.index = index
        self.print_name = clean_name(self.full_name)
        self.avatar = image
    
    def __eq__(self, other):
        return self.full_name == other.full_name
    
class PanelInfo:
    type_font = ImageFont.truetype("Fonts/Montserrat-Regular.ttf", size = 24)
    total_font = ImageFont.truetype("Fonts/Montserrat-Regular.ttf", size = 52)
    rank_font = ImageFont.truetype("Fonts/Montserrat-Regular.ttf", size = 72.5)
    video_font = ImageFont.truetype("Fonts/Montserrat-Regular.ttf", size = 30)

    def __init__(self, info_dict: dict, people: list[Person], save_path: str):
        self.info_dict = info_dict
        self.people = people
        self.base_path = save_path
        # self.background = Image.new("RGBA", (1920, 1080), "black")
        self.background = Image.open(BG_PATH)
        self.video_frame = Image.open(FRAME_PATH)
        self.avatars = len(people)

        self.rows = int(math.ceil(self.avatars / (NUM_COLS_PER_SIDE * 2)))

    def get_offset(self) -> tuple[int, int]:
        return tuple(np.subtract(self.background.size, self.video_frame.size) // 2)
    
    def get_frame_pos(self) -> list[int, int, int, int]:
        offset = self.get_offset()
        vf_w, vf_h = self.video_frame.size
        return [offset[0], offset[0] + vf_w, offset[1], offset[1] + vf_h]

    def calculate_half(
        self, 
        space: tuple[int, int, int, int],
        count: int
    ) -> Generator[tuple[int, int]]:
        left, right, top, bottom = space
        num_full_rows = count // NUM_COLS_PER_SIDE
        remainder = count % NUM_COLS_PER_SIDE
        vertical_size = self.rows * AVATAR_SIZE + Y_PAD * (self.rows - 1)
        y_edge_padding = MIN_EDGE_PADDING / 2 if vertical_size  + MIN_EDGE_PADDING == bottom else (bottom - vertical_size) / 2

        xpos = np.tile(
            np.arange(NUM_COLS_PER_SIDE) * (AVATAR_SIZE + X_PAD) + (left + X_PAD),
            num_full_rows
        )
        num_per_row = np.repeat(NUM_COLS_PER_SIDE, num_full_rows)
        # Leftover
        if remainder > 0:
            extra_padding = ((right - left) - (remainder * AVATAR_SIZE + (remainder - 1) * X_PAD)) / 2
            last_row = np.arange(remainder) * (X_PAD + AVATAR_SIZE) + (left + extra_padding)
            xpos = np.append(xpos, last_row)
            num_per_row = np.append(num_per_row, remainder)
        ypos = np.repeat(
            np.arange(self.rows) * (AVATAR_SIZE + Y_PAD) + y_edge_padding,
            num_per_row
        )
        for i, j in zip(xpos, ypos):
            yield (int(i), int(j))

def create_template(template_details: PanelInfo) -> Image.Image:
    template_image = Image.new('RGBA', template_details.background.size)
    template_image.paste(template_details.background)
    template_image.paste(template_details.video_frame, template_details.get_offset())

    return template_image

def create_song_panel(
    row: tuple, 
    template_details: PanelInfo, 
    template_panel: Image.Image, 
    avatar_positions: list
) -> None:
    info_dict = template_details.info_dict
    people = template_details.people
    panel = template_panel.copy()
    song_info = {
        "song_name": row[info_dict["song_column"]],
        "anime_name": row[info_dict["anime_column"]],
        "song_type": row[info_dict["type_column"]] if info_dict["type_column"] else None,
        "rank": row[info_dict["rank_column"]],
        "total": row[info_dict["total_column"]]
    }
    print(f"[INFO] Creating panel for song {song_info['song_name']}")
    if info_dict["nominator_column"] is None:
        nominator = None
        print(f"[WARNING] No nominator for song {song_info['song_name']} from show {song_info['anime_name']}")
    else:
        nominator = row[info_dict["nominator_column"]]

    write_song_info(song_info, panel, template_details)
    write_user_info(row, panel, people, avatar_positions, nominator)

    save_path = os.path.join(template_details.base_path, f"panel_{song_info['rank']}.png")
    panel.save(save_path)

def write_song_info(
    song_info: dict, 
    panel: Image.Image, 
    template_details: PanelInfo
) -> None:
    offset = template_details.get_offset()
    bg_w, bg_h = template_details.background.size
    vf_w, vf_h = template_details.video_frame.size

    song_name = song_info["song_name"]
    anime_name = song_info["anime_name"]
    rank = str(song_info["rank"])
    total = str(song_info["total"])

    draw = ImageDraw.Draw(panel)
    video_font = template_details.video_font
    total_font = template_details.total_font
    rank_font = template_details.rank_font

    draw.rectangle(((offset), (offset[0] + vf_w, offset[1] + vf_h)), outline=(193, 193, 193), width=1)
    draw.text((bg_w / 2, offset[1] / 2), song_name, font = video_font, fill = 'white', stroke_width=1, stroke_fill='black', anchor="mm")
    draw.text((bg_w / 2, (bg_h + offset[1] + vf_h) / 2), anime_name, font = video_font, fill = 'white', stroke_width=1, stroke_fill='black', anchor="mm")

    rank_box_center = (offset[0] + (RANK_POSITION[0][0] * vf_w + RANK_POSITION[1][0] * vf_w) / 2, 
                       offset[1] + (RANK_POSITION[0][1] * vf_h + RANK_POSITION[1][1] * vf_h) / 2)
    total_box_center = (offset[0] + (TOTAL_POSITION[0][0] * vf_w + TOTAL_POSITION[1][0] * vf_w) / 2, 
                        offset[1] + (TOTAL_POSITION[0][1] * vf_h + TOTAL_POSITION[1][1] * vf_h) / 2)
    type_box_center = (offset[0] + (TYPE_POSITION[0][0] * vf_w + TYPE_POSITION[1][0] * vf_w) / 2, 
                       offset[1] + (TYPE_POSITION[0][1] * vf_h + TYPE_POSITION[1][1] * vf_h) / 2)

    draw.text((rank_box_center), rank, font = rank_font, fill = (244, 186, 23), stroke_width=1, stroke_fill='black', anchor="mm")
    draw.text((total_box_center), total, font = total_font, fill = (244, 186, 23), stroke_width=1, stroke_fill='black', anchor="mm")
    if song_info["song_type"]:
        type  = song_info["song_type"]
        type_font = template_details.type_font
        draw.text((type_box_center), type, font = type_font, fill = (244, 186, 23), stroke_width=1, stroke_fill='black', anchor="mm")

def write_user_info(
    row: tuple, 
    panel: Image.Image, 
    people: list[Person], 
    avatar_positions: list[(int, int)], 
    nominator: str
) -> None:
    low_people, high_people = getLowHighPeople(row, people, nominator)
    
    draw = ImageDraw.Draw(panel)
    name_font = ImageFont.truetype("Fonts/antipasto.regular.ttf", size = 30)
    score_font = ImageFont.truetype("Fonts/SEANSBU.TTF", size = 36)
    for index, person in enumerate(people):
        name = person.full_name
        print_name = person.print_name
        score = row[person.index]
        start_x = avatar_positions[index][0]
        start_y = avatar_positions[index][1]

        text_color = "white"
        if score < 4:
            if score == 1:
                border_color = (253, 255, 114)
            elif score == 2:
                border_color = (229, 229, 229)
            else:
                border_color = (186, 137, 95)
            
            border_size = 10
            glow = create_glow(border_size)
            draw_glow(panel, (start_x - border_size, start_y - border_size), glow, border_color)

        if person in low_people:
            text_color = (53, 182, 31, 255)
        elif person in high_people:
            text_color = (255, 0, 0, 255)
        elif name == nominator:
            text_color = (0, 255, 255, 255)
        
        panel.paste(person.avatar, avatar_positions[index])
        draw.rectangle(((start_x, start_y), (start_x + AVATAR_SIZE, start_y + AVATAR_SIZE)), outline=(193, 193, 193), width=1)
        draw.text((start_x + (AVATAR_SIZE / 2), start_y), print_name, font = name_font, fill = "white", stroke_width=2, stroke_fill='black', anchor="mm")
        draw.text((start_x + (AVATAR_SIZE / 2), start_y + AVATAR_SIZE - score_font.size * .2), str(score), font = score_font, fill = text_color, stroke_width=2, stroke_fill='black', anchor="mm")

def create_glow(glow_width: int) -> Image.Image:
    width = AVATAR_SIZE + glow_width * 2
    def calc_alpha(x, y):
        d = np.abs(np.stack([(x - width / 2), (y - width / 2)])) - AVATAR_SIZE / 2
        l = np.linalg.norm(np.maximum(d, 0), axis=0)
        # m = np.minimum(np.max(d, axis = 0), 0)
        distance = l
        alpha = np.maximum(1.0 - 1.0 / glow_width * distance, 0)
        return (alpha * 255).astype(np.uint8)
    data = np.fromfunction(calc_alpha, shape=(width, width), dtype=np.uint8)
    im = Image.fromarray(data, "L")
    return im

def draw_glow(
    panel: Image.Image, 
    pos: tuple[int, int], 
    glow: Image.Image,
    glow_color = "yellow",
) -> None:
    glow_base = Image.new(mode = "L", size = panel.size)
    x, y = pos
    tmp = Image.new(mode = "L", size = panel.size)
    tmp.paste(glow, (x, y))
    # We don't use Image.paste here because we want to just take the whiter value
    glow_base = ImageChops.lighter(glow_base, tmp)
    # in theory we could paste the glows with a mask on a separate layer with the boxes
    new_im = Image.composite(
        Image.new(mode = "RGBA", size = panel.size, color = glow_color), 
        panel, 
        glow_base,
    )
    panel.paste(new_im)


def clean_name(name: str) -> str:
    name = re.sub(r'\d+', '', name)
    font = ImageFont.truetype("Fonts/antipasto.regular.ttf", size = 30)
    name_length = font.getlength(name)
    while (name_length > AVATAR_SIZE + X_PAD * 2 - 5):
        decrease = AVATAR_SIZE / name_length
        num_char = len(name)
        name = name[:int(num_char * decrease)]
        name_length = font.getlength(name)
    return name

def getLowHighPeople(
    row: tuple,
    people: list[Person], 
    nominator: 
str) -> tuple[list[Person], list[Person]]:
    low_score = None
    high_score = None
    low_people = []
    high_people = []
    for person in people:
        if (person.full_name != nominator):
            index = person.index
            score = row[index]
            if not low_score or score < low_score:
                low_score = score
            elif not high_score or score > high_score:
                high_score = score

    for person in people:
        if (person.full_name != nominator):
            index = person.index
            score = row[index]
            if score == low_score:
                low_people.append(person)
            elif score == high_score:
                high_people.append(person)

    return (low_people, high_people)

def adjust_frame(template_details: PanelInfo, type_column: Optional[int]) -> Image.Image:
    cols = NUM_COLS_PER_SIDE * 2
    bg_w, _ = template_details.background.size
    _, vf_h = template_details.video_frame.size
    avatar_space = (cols + 2) * X_PAD + (cols * AVATAR_SIZE)
    new_width = bg_w - avatar_space
    new_frame =  template_details.video_frame.resize((new_width, vf_h))
    if (type_column is None):
        rect_width = int((TYPE_POSITION[1][0] - TYPE_POSITION[0][0]) * new_width * 1.3)
        rect_height = int((TYPE_POSITION[1][1] - TYPE_POSITION[0][1]) * vf_h * 1.3)
        rect_pos = (new_width - rect_width), 0
        
        transparent = Image.new('RGBA', (rect_width, rect_height), (255, 0, 0, 0))

        new_frame.paste(transparent, rect_pos)
    return new_frame

def get_columns(sheet: openpyxl.worksheet.worksheet.Worksheet) -> tuple[dict, list[Person]]:
    info_dict = {
        "anime_column": None,
        "song_column": None,
        "type_column": None,
        "rank_column": None,
        "total_column": None,
        "nominator_column": None
    }
    people = []
    for index, cell in enumerate(sheet[1]):
        column = cell.value
        if column:
            if "anime" in column.lower():
                info_dict["anime_column"] = index
            elif column.lower() in ['song info', 'songinfo', 'songartist', "song name", "songname"] and info_dict["song_column"] is None:
                info_dict["song_column"] = index
            elif "type" in column.lower():
                info_dict["type_column"] = index
            elif "rank" in column.lower():
                info_dict["rank_column"] = index
            elif "total" in column.lower():
                info_dict["total_column"] = index
            elif "nominator" in column.lower():
                info_dict["nominator_column"] = index
            elif info_dict["total_column"]:
                try:
                    path = os.path.join(AVATAR_PATH, f"{column}.png")
                    avatar = Image.open(path).resize((AVATAR_SIZE, AVATAR_SIZE))
                    people.append(Person(column, index, avatar))
                except:
                    print(f"[WARNING] Could not find image for {column}")
                    img = Image.new('RGBA', (AVATAR_SIZE, AVATAR_SIZE), (0, 0, 0, 255))
                    people.append(Person(column, index, img))
                    
    return (info_dict, people)

def create_all_panels(sheet_name: str, save_path: str) -> None:
    wrkbk  = openpyxl.load_workbook(sheet_name)
    sheet = wrkbk.active
    indices_info, people = get_columns(sheet)

    template_details = PanelInfo(indices_info, people, save_path)
    template_details.video_frame = adjust_frame(template_details, indices_info["type_column"])

    avatar_positions = get_avatar_positions(template_details)
    template_panel = create_template(template_details)
    for index, row in enumerate(sheet.iter_rows(min_row = 2, values_only = True)):
        if row is None or row[0] is None:
            print(f"[INFO] Hit none on row {index + 2}. Exiting")
            break
        create_song_panel(row, template_details, template_panel, avatar_positions)

def get_avatar_positions(template_details: PanelInfo) -> list[tuple[int, int]]:
    frame_pos = template_details.get_frame_pos()
    width, height = template_details.background.size
    avatar_positions_left = template_details.calculate_half(
        space=(0, frame_pos[0], 0, height), 
        count = template_details.avatars // 2,
    )
    avatar_positions_right = template_details.calculate_half(
        space=(frame_pos[1], width, 0, height), 
        count = template_details.avatars - template_details.avatars // 2,
    )
    return list(itertools.chain(avatar_positions_left, avatar_positions_right))

def create_dirs(sheet_name: str) -> str:
    save_path = os.path.join(SAVE_PATH, "".join(sheet_name.split(".")[:-1]), "panels")
    if not os.path.exists(save_path):
        os.makedirs(save_path)
    return save_path


def main(sheet_name: str) -> None:
    save_path = create_dirs(sheet_name)
    file_path = os.path.join(os.getcwd(), sheet_name)
    create_all_panels(file_path, save_path)

if __name__ == '__main__':
    import sys
    try:
        sheet_name  = sys.argv[1]
    except IndexError:
        print(DOC_STRING)
        exit()
    main(sheet_name)