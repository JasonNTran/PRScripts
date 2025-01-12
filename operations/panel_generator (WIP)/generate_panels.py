from PIL import Image, ImageDraw, ImageFont
import openpyxl
import os

BG_PATH = './panel_background.png'
FRAME_PATH = './frame_cropped.png'
AVATAR_PATH = './avatars'
HORIZONTAL_PADDING = 25
VERTICAL_PADDING = 20
VERTICAL_EDGE_PAD = 40
AVATAR_SIZE = 100
MAX_AVATAR_SIZE = 360 # 1/3 of 1080. Useful in the case theres only a handful of people
# Where each box in the frame starts and ends based on pixel percentage. Hardcoded for this template. NYE, Needs topleft and bot right positions
RANK_SPACING = [0, 0] 
TOTAL_SPACING = [0, 0]
TYPE_SPACING = [0, 0]

class TemplateImageDetails:
    def __init__(self, info_dict, name_dict):
        self.info_dict = info_dict
        self.name_dict = name_dict
        self.background = Image.open(BG_PATH)
        self.video_frame = Image.open(FRAME_PATH)
        self.avatars = len(name_dict)

    def get_offset(self):
        bg_w, bg_h = self.background.size
        vf_w, vf_h = self.video_frame.size
        return (bg_w - vf_w)// 2, (bg_h - vf_h) // 2
    
    def get_cols(self):
        rows = (1080 - VERTICAL_EDGE_PAD + VERTICAL_PADDING) // (AVATAR_SIZE + VERTICAL_PADDING)
        return self.avatars // rows
    
# Returns a template and all the necessary reuseable variables for creating each panel
def create_template(template_details, avatar_positions):
    template_image = Image.new('RGBA', template_details.background.size)
    template_image.paste(template_details.background)
    template_image.paste(template_details.video_frame, template_details.get_offset())

    for index, name in enumerate(template_details.name_dict):
        path = os.path.join(AVATAR_PATH, f"{name}.png")
        icon = Image.open(path).resize((AVATAR_SIZE, AVATAR_SIZE))
        template_image.paste(icon, avatar_positions[index])
    template_image.save("example_temaplate_avatars.png") # for debugging purposes
    # template_image.show()
    return template_image

def calculate_avatar_positions(template_details):
    positions = []
    cols = template_details.get_cols() // 2
    avatars = template_details.avatars
    vf_w, _ = template_details.video_frame.size

    for i in range(0, avatars):
        if i >= avatars // 2:
            r = (i - (avatars // 2)) // cols
            c = i % cols
            horizontaL_offset = vf_w + template_details.get_offset()[0]
            vertical_pos = (r * AVATAR_SIZE) + (r + 1) * HORIZONTAL_PADDING
            horizontal_pos = (c * AVATAR_SIZE) + (c + 1) * VERTICAL_PADDING + horizontaL_offset
        else:
            r = i // cols
            c = i % cols
            vertical_pos = (r * AVATAR_SIZE) + (r + 1) * HORIZONTAL_PADDING
            horizontal_pos = (c * AVATAR_SIZE) + (c + 1) * VERTICAL_PADDING
        positions.append((horizontal_pos, vertical_pos))
    return positions

def create_all_panels(sheet_name):
    wrkbk  = openpyxl.load_workbook(sheet_name)
    sheet = wrkbk.active
    info_dict, name_dict = get_columns(sheet)
    template_details = TemplateImageDetails(info_dict, name_dict)
    resize_frame(template_details)
    avatar_positions = calculate_avatar_positions(template_details)
    template_panel = create_template(template_details, avatar_positions)
    for row in sheet.iter_rows(min_row = 2):
        create_song_panel(row, template_details, template_panel, avatar_positions)

def create_song_panel(row, template_details, template_panel, avatar_positions):
    info_dict = template_details.info_dict
    name_dict = template_details.name_dict
    panel = template_panel.copy()
    song_info = {
        "song_name": row[info_dict["song_column"]].value,
        "anime_name": row[info_dict["anime_column"]].value,
        # "song_type": row[info_dict["type_column"]].valie,
        "rank": row[info_dict["rank_column"]].value,
        "total": row[info_dict["total_column"]].value
    }

    write_song_info(song_info, panel, template_details)
    write_avatar_info(row, name_dict, panel, avatar_positions)

    save_path = os.path.join("test", "panels", f"panel_{song_info["rank"]}.png")
    panel.save(save_path)
    # return None

def write_song_info(song_info, panel, template_details):
    offset = template_details.get_offset()
    bg_w, bg_h = template_details.background.size
    _, vf_h = template_details.video_frame.size

    song_name = song_info["song_name"]
    anime_name = song_info["anime_name"]
    rank = song_info["rank"]

    draw = ImageDraw.Draw(panel)
    font_size = offset[1] // 1.5
    font = ImageFont.truetype("antipasto-demibold.ttf", font_size)
    # Song name on top
    _, _, w, _ = draw.textbbox((0, 0), song_name, font = font)
    draw.text(((bg_w - w) // 2, offset[1] - font_size - 5), song_name, font = font, fill = 'white')

    # Anime name on bottom
    _, _, w, _ = draw.textbbox((0, offset[1] + vf_h), anime_name, font = font)
    draw.text(((bg_w - w) // 2, (offset[1] + vf_h) + 5), anime_name, font = font, fill = 'white')

    # Next we need to put next for rank, total, and type
 
def write_avatar_info(row, name_dict, panel, avatar_positions):
    lowName, highName = getLowHighIndex(row, name_dict)

    draw = ImageDraw.Draw(panel)
    font_size = 30
    font = ImageFont.truetype("antipasto-demibold.ttf", size = font_size)
    nominator = "shiroky"
    for index, name in enumerate(name_dict):
        score = str(row[name_dict[name]].value)
        start_x = avatar_positions[index][0]
        start_y = avatar_positions[index][1]
        name = name
        name_length = draw.textlength(name, font = font)
        score_length = draw.textlength(score, font = font)
        print("[INFO] Writing avatar info", flush=True)
        color = "white"
        if name == highName:
            color = (0, 255, 0, 128)
        elif name == lowName:
            color = (255, 0, 0, 128)
        elif name == nominator:
            color = (0, 0, 255, 128)
        draw.rectangle(((start_x, start_y ), (start_x + AVATAR_SIZE, start_y + AVATAR_SIZE)), outline=color, width=1)
        draw.text((start_x + (AVATAR_SIZE - name_length) / 2, start_y - (font_size / 2)), name, font = font, fill = color)
        draw.text((start_x + (AVATAR_SIZE - score_length) / 2, start_y + (AVATAR_SIZE - font_size / 2)), score, font = font, fill = color)
def resize_name(name):
    print("[INFO] Resizing name", flush=True)
    font = ImageFont.truetype("antipasto-demibold.ttf", size = 30)
    name_length = font.getlength(name)
    while (name_length > AVATAR_SIZE):
        decrease = AVATAR_SIZE / name_length
        num_char = len(name)
        name = f"{name[:int(num_char * decrease)]}..."
        name_length = font.getlength(name)
    return name

def getLowHighIndex(row, name_dict):
    print("[INFO] Getting highest and lowest values for song", flush=True)
    low = None
    high = None
    for name, column_index in name_dict.items():
        score = row[column_index].value
        if not low or score < row[name_dict[low]].value:
            low = name
        elif not high or score > row[name_dict[high]].value:
            high = name
    return low, high
def create_dirs():
    save_path = os.path.join("test", "panels")
    if not os.path.exists(save_path):
        os.makedirs(save_path)

def resize_frame(template_details):
    cols = template_details.get_cols()
    bg_w, _ = template_details.background.size
    _, vf_h = template_details.video_frame.size
    avatar_space = (cols + 1) * HORIZONTAL_PADDING + (cols * AVATAR_SIZE)
    new_width = bg_w - avatar_space
    template_details = template_details.video_frame.resize((new_width, vf_h))
    # get at what pixel percent the bubbles appear on
    # math is pos = vf_h * %, vf_w * %
def get_columns(sheet):
    info_dict = {
        "anime_column": None,
        "song_column": None,
        "link_column": None,
        "type_column": None,
        "rank_column": None,
        "total_column": None
    }
    name_dict = {}
    for index, cell in enumerate(sheet[1]):
        column = cell.value
        if column:
            if "anime" in column.lower():
                info_dict["anime_column"] = index
            elif column.lower() in ['song info', 'songinfo', 'songartist'] and info_dict["song_column"] is None:
                info_dict["song_column"] = index
            elif column.lower() in ['song info', 'songinfo']:
                info_dict["link_column"] = index
            elif "type" in column.lower():
                info_dict["type_column"] = index
            elif "rank" in column.lower():
                info_dict["rank_column"] = index
            elif "total" in column.lower():
                info_dict["total_column"] = index
            elif info_dict["total_column"]:
                name_dict[column] = index
    if info_dict["link_column"] == None:
        info_dict["link_column"] = info_dict["song_column"] + 1 
    
    return (info_dict, name_dict)
create_dirs()
create_all_panels("test.xlsx")
