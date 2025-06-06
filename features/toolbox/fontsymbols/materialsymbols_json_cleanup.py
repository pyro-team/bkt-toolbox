import json
import os.path

# Load the JSON file
input_file = "materialsymbols.json"
output_file = "materialsymbols_cleaned.json"

with open(os.path.join(os.path.dirname(os.path.realpath(__file__)), input_file), "r", encoding="utf-8") as file:
    data = json.load(file)

font_names = [
    "Material Symbols Outlined",
    "Material Symbols Rounded",
    "Material Symbols Sharp"
    ]

font_names_legacy = [
    "Material Icons",
    "Material Icons Outlined",
    "Material Icons Round",
    "Material Icons Sharp",
    "Material Icons Two Tone"
]

# Iterate over the icons list and filter out unwanted icons
filtered_icons = []
for icon in data['icons']:
    unsupported_font = icon['unsupported_families']
    if not all(font in unsupported_font for font in font_names):
        # Keep the icon if it is supported in any of the specified fonts
        icon['unsupported_families'] = [
            font for font in unsupported_font if font not in font_names_legacy
        ]
        filtered_icons.append(icon)

# Update the icons list with the filtered icons
data['icons'] = filtered_icons

# Save the cleaned JSON to a new file
with open(os.path.join(os.path.dirname(os.path.realpath(__file__)), output_file), "w", encoding="utf-8") as file:
    json.dump(data, file, ensure_ascii=False, indent=2)

print(f"Cleaned JSON saved to {output_file}")