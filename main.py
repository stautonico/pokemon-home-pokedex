import requests
from pprint import pprint
import os
import json
import xlsxwriter

API_BASE = "https://pokeapi.co/api/v2"


def download_all_sprites():
    r = requests.get(f"{API_BASE}/pokemon?limit=100000&offset=0")

    for pokemon in r.json()["results"]:
        r2 = requests.get(pokemon["url"])

        if r2.status_code != 200:
            print(f"[FAIL] Failed to download {pokemon['name']}")

            with open("sprites-failed.txt", "a") as f:
                f.write(f"{pokemon['name']}\n")
                continue

        print(f"[INFO] Downloading sprite for {r2.json()['name']}")

        # Check if we already have the sprite
        if os.path.exists(f"sprites/{r2.json()['name']}.png"):
            print(f"[INFO] Already have sprite for {r2.json()['name']}")
            continue

        image_url = r2.json()["sprites"]["front_default"]

        if image_url is None:
            print(f"[INFO] No sprite for {r2.json()['name']}")
            continue

        # Download the image url into the sprites folder
        r = requests.get(image_url)

        if r.status_code != 200:
            with open("sprites-failed.txt", "a") as f:
                f.write(f"{pokemon['name']}\n")
                continue

        with open(f"sprites/{r2.json()['name']}.png", "wb") as f:
            f.write(r.content)

        print(f"[DONE] Downloaded sprite for {r2.json()['name']}")


def make_boxes():
    with open("boxes.json", "r") as f:
        boxes_json = json.loads(f.read())

    workbook = xlsxwriter.Workbook('pokemon.xlsx')

    worksheet = workbook.add_worksheet()

    # Boxes are 6x5
    # Boxes are separated by 1 row
    # Boxes also have one merged row on top that contains its title

    # The "boxes" in the json file are an object
    # They contain the following keys:
    # "title" - The title of the box (str)
    # "pokemon" - The pokemon in the box (list of strings) (should be 30 pokemon, but can be less)
    # Some of the pokemon in the boxes are dictionaries (the gmax ones), we just want the "pid" key

    current_row = 0

    for box in boxes_json["boxes"]:
        # Add the title
        worksheet.merge_range(current_row, 0, current_row, 5, box["title"], workbook.add_format({'bold': True}))
        current_row += 1

        # Add the pokemon
        for i, pokemon in enumerate(box["pokemon"]):
            # Get the pokemon id
            if isinstance(pokemon, dict):
                pokemon_id = pokemon["pid"]
            else:
                pokemon_id = pokemon

            # Get the pokemon name (we'll just use the id for now)
            pokemon_name = pokemon_id
            # r = requests.get(f"{API_BASE}/pokemon/{pokemon_id}")
            # pokemon_name = r.json()["name"]

            # Get the pokemon sprite
            if os.path.exists(f"sprites/{pokemon_name}.png"):
                # Write the google sheets image function (and get from http://localhost:8000)
                worksheet.write(current_row, i % 6, f'=IMAGE("http://localhost:8000/sprites/{pokemon_name}.png", 4)')

            else:
                worksheet.write(current_row, i % 6, pokemon_name)

            if i % 6 == 5:
                current_row += 1

        current_row += 2

    workbook.close()


if __name__ == '__main__':
    make_boxes()
