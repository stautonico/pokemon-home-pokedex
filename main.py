import requests
import os
import json
import xlsxwriter

API_BASE = "https://pokeapi.co/api/v2"

GITHUB_SPRITE_URL = "https://raw.githubusercontent.com/stautonico/pokemon-home-pokedex/main/sprites"

workbook = xlsxwriter.Workbook('pokemon.xlsx')

checklist = workbook.add_worksheet("Checklist")
boxes = workbook.add_worksheet("Boxes")

pokemon_cells = {}

center_text = workbook.add_format()
center_text.set_align("center")
center_text.set_align("vcenter")

class Colors:
    INFO = '\033[94m'
    SUCCESS = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    END = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

def info(msg):
    print(f"{Colors.INFO}[INFO]{Colors.END}: {msg}")

def success(msg):
    print(f"{Colors.SUCCESS}[GOOD]{Colors.END}: {msg}")

def warn(msg):
    print(f"{Colors.WARNING}[WARN]{Colors.END}: {msg}")

def fail(msg):
    print(f"{Colors.FAIL}[FAIL]: {msg}{Colors.END}")


# TODO: Add where to get them (which game, etc)
# TODO: Add sprites for additional forms (e.g. regional, variants, gmax, m/f, etc)

# Manual file renames:
# mime-jr -> mimejr
# mr-mime -> mrmime
# ho-oh -> hooh
# cp deoxys-normal -> deoxys
# cp wormadam-plant -> wormadam
# cp giratina-altered -> giratina
# cp shaymin-land -> shaymin
# cp basculin-red-striped -> basculin
# cp darmanitan-standard -> darmanitan
# cp tornadus-incarnate -> tornadus
# cp thundurus-incarnate -> thundurus
# cp landorus-incarnate -> landorus
# cp keldeo-ordinary -> keldeo
# cp meloetta-aria -> meloetta
# cp meowstic-male -> meowstic
# cp aegislash-shield -> aegislash
# cp pumpkaboo-average -> pumpkaboo
# cp gourgeist-average -> gourgeist
# cp zygarde-50 -> zygarde
# cp oricorio-pom-pom -> oricorio
# cp lycanroc-midday -> lycanroc
# cp wishiwashi-solo -> wishiwashi
# type-null -> typenull
# cp mimikyu-disguised -> mimikyu
# cp toxtricity-amped -> toxtricity
# mr-rime -> mrrime
# cp eiscue-ice -> eiscue
# cp indeedee-male -> indeedee
# cp urshifu-single-strike -> urshifu
# cp basculegion-male -> basculegion
# cp enamorus-incarnate -> enamorus

def download_all_sprites():
    r = requests.get(f"{API_BASE}/pokemon?limit=100000&offset=0")

    for pokemon in r.json()["results"]:
        r2 = requests.get(pokemon["url"])

        if r2.status_code != 200:
            fail(f"Failed to download {pokemon['name']}")

            with open("sprites-failed.txt", "a") as f:
                f.write(f"{pokemon['name']}\n")
                continue

        info(f"Downloading sprite for {r2.json()['name']}")

        # Check if we already have the sprite
        if os.path.exists(f"sprites/{r2.json()['name']}.png"):
            info(f"Already have sprite for {r2.json()['name']}")
            continue

        image_url = r2.json()["sprites"]["front_default"]

        if image_url is None:
            warn(f"No sprite for {r2.json()['name']}")
            continue

        # Download the image url into the sprites folder
        r = requests.get(image_url)

        if r.status_code != 200:
            with open("sprites-failed.txt", "a") as f:
                f.write(f"{pokemon['name']}\n")
                continue

        with open(f"sprites/{r2.json()['name']}.png", "wb") as f:
            f.write(r.content)

        success(f"Downloaded sprite for {r2.json()['name']}")


def make_checklist():
    checklist.write(0, 0, "Caught")
    checklist.write(0, 1, "ID")
    checklist.write(0, 2, "Name")
    checklist.write(0, 3, "Sprite")

    # Write the totals (I2, J2, and I3, J3)
    checklist.write(1, 8, "Species", center_text)
    checklist.write(1, 9, "=concat(countif(A2:A906, \"TRUE\"), concat(\" of \", COUNTA(A2:A906)))", center_text)
    checklist.write(1, 10, "=concat(round(countif(A2:A906, \"TRUE\")/COUNTA(A2:A906)*100, 2), \"%\")", center_text)
    # put a border around the totals
    checklist.conditional_format(1, 8, 1, 8, {"type": "no_blanks",
                                              "format": workbook.add_format({"left": 1, "top": 1, "bottom": 1})})
    checklist.conditional_format(1, 9, 1, 9, {"type": "no_blanks",
                                              "format": workbook.add_format({"top": 1, "bottom": 1})})
    checklist.conditional_format(1, 10, 1, 10, {"type": "no_blanks",
                                                "format": workbook.add_format({"right": 1, "top": 1, "bottom": 1})})

    checklist.conditional_format(2, 8, 2, 8, {"type": "no_blanks",
                                              "format": workbook.add_format({"left": 1, "top": 1, "bottom": 1})})
    checklist.conditional_format(2, 9, 2, 9, {"type": "no_blanks",
                                              "format": workbook.add_format({"top": 1, "bottom": 1})})
    checklist.conditional_format(2, 10, 2, 10, {"type": "no_blanks",
                                                "format": workbook.add_format({"right": 1, "top": 1, "bottom": 1})})

    checklist.write(2, 8, "Total", center_text)
    checklist.write(2, 9, "=concat(countif(A2:A1293, \"TRUE\"), concat(\" of \", COUNTA(A2:A1293)))", center_text)
    checklist.write(2, 10, "=concat(round(countif(A2:A1293, \"TRUE\")/COUNTA(A2:A1293)*100, 2), \"%\")", center_text)

    row = 1

    with open("all-pokemon.json", "r") as f:
        all_pokemon = json.loads(f.read())["pokemon"]

    for pokemon in all_pokemon:
        gmax = False
        # Special case for gmax
        if pokemon.startswith("gmax-"):
            # Read the pokemon data without the "gmax-" prefix
            pokemon = pokemon[5:]
            gmax = True

        # Load the pokemon data (from the json file)
        try:
            with open(f"pokemon-data/{pokemon}.json", "r") as f:
                pokemon_data = json.loads(f.read())
        except FileNotFoundError:
            warn(f"Could not find '{pokemon}', using '{pokemon}' as the name")
            pokemon_data = {"name": pokemon, "id": 0}

        checklist.write(row, 0, "FALSE", center_text)
        # Write the ID
        checklist.write(row, 1, pokemon_data["id"], center_text)
        # Write the name
        if gmax:
            checklist.write(row, 2, f"Gigantamax {pokemon_data['name'].title()}", center_text)
        else:
            checklist.write(row, 2, pokemon_data["name"].title(), center_text)
        # Write the sprite
        if os.path.exists(f"sprites/{pokemon_data['name']}.png"):
            # Add the image (using google sheets image url)
            checklist.write(row, 3, f'=IMAGE("{GITHUB_SPRITE_URL}/{pokemon_data["name"]}.png", 2)')
        else:
            checklist.write(row, 3, f"TODO: {pokemon_data['name']} (image not found)")

        # For columns 0-3, if the first column in the row is "TRUE", make the background green
        # otherwise, make the background red

        checklist.conditional_format(row, 0, row, 3, {
            "type": "formula",
            "criteria": f'=COUNTIF($A{row + 1}, "TRUE") = 1',
            "format": workbook.add_format({"bg_color": "#00FF00"})
        })

        checklist.conditional_format(row, 0, row, 3, {
            "type": "formula",
            "criteria": f'=COUNTIF($A{row + 1}, "FALSE") = 1',
            "format": workbook.add_format(
                {"bg_color": "#FF0000", "font_color": "#FFFFFF"})
        })

        # Add a border to each row
        checklist.conditional_format(row, 0, row, 3, {
            "type": "no_blanks",
            "format": workbook.add_format({"bottom": 1, "top": 1})
        })

        checklist.conditional_format(row, 0, row, 0, {
            "type": "no_blanks",
            "format": workbook.add_format({"left": 1})
        })

        checklist.conditional_format(row, 3, row, 3, {
            "type": "no_blanks",
            "format": workbook.add_format({"right": 1})
        })

        # Make the sprite column 96px wide
        checklist.set_column_pixels(3, 3, 96)
        checklist.set_row_pixels(row, 96)

        # Set the position of the checkbox
        pokemon_cells[pokemon] = "A" + str(row + 1)

        row += 1


def make_boxes():
    with open("boxes.json", "r") as f:
        boxes_json = json.loads(f.read())

    # The "boxes" in the json file are an object
    # They contain the following keys:
    # "title" - The title of the box (str)
    # "pokemon" - The pokemon in the box (list of strings) (should be 30 pokemon, but can be less)
    # Some of the pokemon in the boxes are dictionaries (the gmax ones), we just want the "pid" key

    row = 1
    col = 1
    box = 0

    while True:
        try:
            pokemon_index = 0
            current_box = boxes_json["boxes"][box]
        except IndexError:
            break  # We're done

        # Merge the first 6 cells in the row
        boxes.merge_range(row, col, row, col + 5, current_box["title"])
        boxes.conditional_format(row, col, row, col + 5, {
            "type": "no_blanks",
            "format": workbook.add_format({"bottom": 1, "top": 1, "left": 1, "right": 1})
        })
        # Center the text and make it bold + 16px
        # boxes.set_row(row, 24)
        # boxes.set_column_pixels(col, col + 5, 24)
        boxes.write(row, col, current_box["title"],
                    workbook.add_format({"bold": True, "font_size": 16, "align": "center", "valign": "vcenter"}))

        row += 1

        for _ in range(5):
            for _ in range(6):
                try:
                    pokemon = current_box["pokemon"][pokemon_index]
                except IndexError:
                    # If we run out of pokemon, just break (this box doesn't have 30 pokemon)
                    break

                pokemon_index += 1

                if isinstance(pokemon, dict):
                    pokemon = pokemon["pid"]

                if os.path.exists(f"sprites/{pokemon}.png"):
                    # Add the image (using google sheets image url)
                    boxes.write(row, col, f'=IMAGE("{GITHUB_SPRITE_URL}/{pokemon}.png", 2)')
                else:
                    boxes.write(row, col, f"TODO: {pokemon} (image not found)")

                checkbox_cell = pokemon_cells.get(pokemon, None)

                boxes.conditional_format(row, col, row, col, {
                    "type": "formula",
                    "criteria": f'=COUNTIF(INDIRECT(\"Checklist!${checkbox_cell}\"), "TRUE") = 1',
                    "format": workbook.add_format({"bg_color": "#00FF00"})
                })

                boxes.conditional_format(row, col, row, col, {
                    "type": "formula",
                    "criteria": f'=COUNTIF(INDIRECT(\"Checklist!${checkbox_cell}\"), "FALSE") = 1',
                    "format": workbook.add_format({"bg_color": "#FF0000", "font_color": "#FFFFFF"})
                })

                # Set a border around each cell
                boxes.conditional_format(row, col, row, col, {
                    "type": "no_blanks",
                    "format": workbook.add_format({"bottom": 1, "top": 1, "left": 1, "right": 1})
                })

                boxes.set_column_pixels(col, col, 96)
                boxes.set_row_pixels(row, 96)

                col += 1
            row += 1
            col = 1

        # Move the row back up 6 rows to do the second box
        row -= 6
        col = 8
        box += 1

        try:
            current_box = boxes_json["boxes"][box]
            pokemon_index = 0
        except IndexError:
            break  # We're done

        # Merge the first 6 cells in the row
        boxes.merge_range(row, col, row, col + 5, current_box["title"])
        boxes.conditional_format(row, col, row, col + 5, {
            "type": "no_blanks",
            "format": workbook.add_format({"bottom": 1, "top": 1, "left": 1, "right": 1})
        })
        # Center the text and make it bold + 16px
        # boxes.set_row_pixels(row, 24)
        # boxes.set_column_pixels(col, col + 5, 24)
        boxes.write(row, col, current_box["title"],
                    workbook.add_format({"bold": True, "font_size": 16, "align": "center", "valign": "vcenter"}))

        row += 1

        for _ in range(5):
            for _ in range(6):
                try:
                    pokemon = current_box["pokemon"][pokemon_index]
                except IndexError:
                    # If we run out of pokemon, just break (this box doesn't have 30 pokemon)
                    break

                pokemon_index += 1

                if isinstance(pokemon, dict):
                    pokemon = pokemon["pid"]

                if os.path.exists(f"sprites/{pokemon}.png"):
                    # Add the image (using google sheets image url)
                    boxes.write(row, col, f'=IMAGE("{GITHUB_SPRITE_URL}/{pokemon}.png", 2)')
                else:
                    boxes.write(row, col, f"TODO: {pokemon} (image not found)")

                checkbox_cell = pokemon_cells.get(pokemon, None)

                checkbox_cell = pokemon_cells.get(pokemon, None)

                boxes.conditional_format(row, col, row, col, {
                    "type": "formula",
                    "criteria": f'=COUNTIF(INDIRECT(\"Checklist!${checkbox_cell}\"), "TRUE") = 1',
                    "format": workbook.add_format({"bg_color": "#00FF00"})
                })

                boxes.conditional_format(row, col, row, col, {
                    "type": "formula",
                    "criteria": f'=COUNTIF(INDIRECT(\"Checklist!${checkbox_cell}\"), "FALSE") = 1',
                    "format": workbook.add_format({"bg_color": "#FF0000", "font_color": "#FFFFFF"})
                })

                # Set a border around each cell
                boxes.conditional_format(row, col, row, col, {
                    "type": "no_blanks",
                    "format": workbook.add_format({"bottom": 1, "top": 1, "left": 1, "right": 1})
                })

                # Set the column width to 96
                boxes.set_column_pixels(col, col, 96)
                boxes.set_row_pixels(row, 96)

                col += 1
            row += 1
            col = 8

        # One for padding and one for the next box
        row += 2
        col = 1
        box += 1


if __name__ == '__main__':
    make_checklist()
    make_boxes()

    workbook.close()
