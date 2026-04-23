"""Theme palettes for the Tkinter interface."""

from __future__ import annotations


THEMES: dict[str, dict[str, str]] = {
    "Midnight": {
        "bg": "#080910",
        "chrome": "#101119",
        "surface": "#161822",
        "surface_2": "#1D202B",
        "border": "#282C3B",
        "text": "#E3E6EF",
        "muted": "#8990A5",
        "accent": "#00B8D9",
        "accent_hover": "#35CFEA",
        "on_accent": "#080910",
        "ok": "#72D083",
        "warn": "#FFD166",
        "error": "#FF7474",
        "selection": "#122430",
    },
    "Graphite": {
        "bg": "#0B0C0F",
        "chrome": "#121418",
        "surface": "#191C22",
        "surface_2": "#22262E",
        "border": "#303641",
        "text": "#E9ECF1",
        "muted": "#98A1AE",
        "accent": "#5BE6A4",
        "accent_hover": "#79F0B8",
        "on_accent": "#07100C",
        "ok": "#67D98A",
        "warn": "#E8C35A",
        "error": "#FF7C7C",
        "selection": "#173226",
    },
    "Ocean": {
        "bg": "#071016",
        "chrome": "#0D1821",
        "surface": "#132230",
        "surface_2": "#1B2B3A",
        "border": "#294052",
        "text": "#E7F2F7",
        "muted": "#91A8B6",
        "accent": "#3DD6C6",
        "accent_hover": "#6CE4D8",
        "on_accent": "#061312",
        "ok": "#7AD68A",
        "warn": "#FFD66E",
        "error": "#FF7A86",
        "selection": "#12363B",
    },
    "White": {
        "bg": "#F7F8FB",
        "chrome": "#FFFFFF",
        "surface": "#F2F5F9",
        "surface_2": "#FFFFFF",
        "border": "#D6DDE8",
        "text": "#18202F",
        "muted": "#5F6C7F",
        "accent": "#007A9A",
        "accent_hover": "#009EC4",
        "on_accent": "#FFFFFF",
        "ok": "#177245",
        "warn": "#9A6A00",
        "error": "#C73844",
        "selection": "#E3F4FA",
    },
}


def get_theme(name: str | None) -> dict[str, str]:
    return THEMES.get(name or "", THEMES["Midnight"])
