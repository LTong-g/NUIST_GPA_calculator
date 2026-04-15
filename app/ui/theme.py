DAY_THEME = {
    "root": "#ffe0ff",
    "words": "black",
    "button_bg": "#ffe190",
    "button_abg": "#efd180",
    "entry_bg": "#f0f0f0",
}

NIGHT_THEME = {
    "root": "#202020",
    "words": "#e0e0e0",
    "button_bg": "#001e7f",
    "button_abg": "#000e6f",
    "entry_bg": "#2f2f2f",
}


def get_theme_colors(skin):
    if skin == "day":
        return DAY_THEME
    return NIGHT_THEME

