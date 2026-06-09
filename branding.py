from pptx.dml.color import RGBColor

# Brand Palette Constants
COLORS = {
    "BLACK": "111111",      # Deep primary background & headers
    "GOLD": "C8A84B",       # Accent lines, titles, and BKFC data highlights
    "SILVER": "BFC5CC",     # Neutral secondary text, borders, and league benchmarks
    "WHITE": "FFFFFF",      # Content area backgrounds
    "DARK_GRAY": "222222",  # Subtle block backgrounds
    "GRID": "E6E6E6"        # Subtle gridlines for tables
}

def get_rgb(hex_str):
    """Converts a hex string directly to a python-pptx RGBColor object."""
    return RGBColor(
        int(hex_str[0:2], 16),
        int(hex_str[2:4], 16),
        int(hex_str[4:6], 16)
    )