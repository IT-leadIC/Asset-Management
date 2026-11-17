from pathlib import Path
from urllib.parse import quote

import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont, ImageChops


# ============================================================
# PROJECT SETTINGS
# ============================================================

PROJECT_DIR = Path(__file__).resolve().parent

EXCEL_FILE = PROJECT_DIR / "IT_Assets.xlsx"
LOGO_FILE = PROJECT_DIR / "logo.png"

OUTPUT_DIR = PROJECT_DIR / "QR_Codes"

# Live Render application
BASE_URL = "https://asset-management-1o2f.onrender.com"

# Excel header is on row 2
# pandas uses zero-based indexing, so row 2 = 1
EXCEL_HEADER_ROW = 1


# ============================================================
# QR GENERATION MODE
# ============================================================

# False = Skip QR files that already exist
# True  = Regenerate/overwrite all QR files
REGENERATE_EXISTING = False

# This file stores only QR files created in the current run
NEW_QR_LIST = PROJECT_DIR / "new_qr_files.txt"


# ============================================================
# LABEL SIZE
# ============================================================

# COMPLETE LABEL = 26 x 26 mm
LABEL_MM = 26

# Print resolution
DPI = 300


def mm_to_px(mm):
    return int(round(mm * DPI / 25.4))


LABEL_SIZE = mm_to_px(LABEL_MM)


# ============================================================
# QR SIZE
# ============================================================

QR_MM = 22.5
QR_SIZE = mm_to_px(QR_MM)


# ============================================================
# ASSET ID SETTINGS
# ============================================================

# Very small gap between QR and Asset ID
TEXT_GAP_MM = 0.15
TEXT_GAP = mm_to_px(TEXT_GAP_MM)

# Asset ID text size
FONT_SIZE_MM = 2.6
FONT_SIZE = mm_to_px(FONT_SIZE_MM)


# ============================================================
# LOGO SETTINGS
# ============================================================

# Width of LeadIC logo inside QR
LOGO_WIDTH_MM = 6.5
LOGO_WIDTH = mm_to_px(LOGO_WIDTH_MM)

# White space around logo
LOGO_PADDING_X_MM = 0.7
LOGO_PADDING_Y_MM = 0.45


# ============================================================
# CREATE OUTPUT FOLDER
# ============================================================

OUTPUT_DIR.mkdir(exist_ok=True)


# ============================================================
# CHECK REQUIRED FILES
# ============================================================

if not EXCEL_FILE.exists():
    raise FileNotFoundError(
        f"Excel file not found:\n{EXCEL_FILE}"
    )

if not LOGO_FILE.exists():
    raise FileNotFoundError(
        f"Logo file not found:\n{LOGO_FILE}"
    )


# ============================================================
# READ EXCEL
# ============================================================

print("Reading Excel...")

df = pd.read_excel(
    EXCEL_FILE,
    header=EXCEL_HEADER_ROW
)

if "Asset No" not in df.columns:
    raise ValueError(
        "Column 'Asset No' was not found in Excel."
    )

print(f"Assets found: {len(df)}")


# ============================================================
# LOAD LOGO
# ============================================================

logo = Image.open(
    LOGO_FILE
).convert("RGBA")


# ============================================================
# REMOVE EXTRA WHITE SPACE FROM LOGO
# ============================================================

rgb_logo = logo.convert("RGB")

white_background = Image.new(
    "RGB",
    rgb_logo.size,
    "white"
)

difference = ImageChops.difference(
    rgb_logo,
    white_background
)

bbox = difference.getbbox()

if bbox:
    logo = logo.crop(bbox)


# ============================================================
# RESIZE LOGO
# KEEP ORIGINAL RECTANGULAR PROPORTION
# ============================================================

original_width, original_height = logo.size

logo_height = int(
    original_height
    * LOGO_WIDTH
    / original_width
)

logo = logo.resize(
    (
        LOGO_WIDTH,
        logo_height
    ),
    Image.Resampling.LANCZOS
)


# ============================================================
# FONT
# ============================================================

FONT_PATH = r"C:\Windows\Fonts\arialbd.ttf"

if not Path(FONT_PATH).exists():
    FONT_PATH = r"C:\Windows\Fonts\calibrib.ttf"

if not Path(FONT_PATH).exists():
    raise FileNotFoundError(
        "Arial Bold or Calibri Bold font was not found."
    )

font = ImageFont.truetype(
    FONT_PATH,
    FONT_SIZE
)


# ============================================================
# TRACK CURRENT RUN
# ============================================================

new_qr_files = []

generated = 0
skipped = 0


# ============================================================
# GENERATE QR LABELS
# ============================================================

for _, row in df.iterrows():

    # --------------------------------------------------------
    # GET ASSET ID
    # --------------------------------------------------------

    asset_id = str(
        row["Asset No"]
    ).strip()

    if not asset_id or asset_id.lower() == "nan":
        continue


    # --------------------------------------------------------
    # OUTPUT FILE
    # --------------------------------------------------------

    output_file = (
        OUTPUT_DIR
        / f"{asset_id}.png"
    )


    # --------------------------------------------------------
    # SKIP EXISTING QR
    # --------------------------------------------------------

    if (
        output_file.exists()
        and not REGENERATE_EXISTING
    ):

        print(
            f"Skipping existing QR: {asset_id}"
        )

        skipped += 1
        continue


    if (
        REGENERATE_EXISTING
        and output_file.exists()
    ):

        print(
            f"Regenerating QR: {asset_id}"
        )

    else:

        print(
            f"Generating new QR: {asset_id}"
        )


    # ========================================================
    # ASSET URL
    # ========================================================

    asset_url = (
        f"{BASE_URL}/asset/{quote(asset_id)}"
    )


    # ========================================================
    # CREATE QR
    # ========================================================

    qr = qrcode.QRCode(
        version=None,

        # Highest error correction
        # because the center contains the logo
        error_correction=qrcode.constants.ERROR_CORRECT_H,

        box_size=10,
        border=3
    )

    qr.add_data(asset_url)
    qr.make(fit=True)

    qr_image = qr.make_image(
        fill_color="black",
        back_color="white"
    ).convert("RGBA")


    # ========================================================
    # RESIZE QR
    # ========================================================

    qr_image = qr_image.resize(
        (
            QR_SIZE,
            QR_SIZE
        ),
        Image.Resampling.NEAREST
    )


    # ========================================================
    # CREATE WHITE CENTER PANEL
    # ========================================================

    pad_x = mm_to_px(
        LOGO_PADDING_X_MM
    )

    pad_y = mm_to_px(
        LOGO_PADDING_Y_MM
    )

    panel_width = (
        logo.width
        + pad_x * 2
    )

    panel_height = (
        logo.height
        + pad_y * 2
    )

    white_panel = Image.new(
        "RGBA",
        (
            panel_width,
            panel_height
        ),
        "white"
    )


    # ========================================================
    # PUT LOGO ON WHITE PANEL
    # ========================================================

    logo_x = (
        panel_width
        - logo.width
    ) // 2

    logo_y = (
        panel_height
        - logo.height
    ) // 2

    white_panel.alpha_composite(
        logo,
        (
            logo_x,
            logo_y
        )
    )


    # ========================================================
    # PUT LOGO PANEL IN QR CENTER
    # ========================================================

    panel_x = (
        QR_SIZE
        - panel_width
    ) // 2

    panel_y = (
        QR_SIZE
        - panel_height
    ) // 2

    qr_image.alpha_composite(
        white_panel,
        (
            panel_x,
            panel_y
        )
    )


    # ========================================================
    # CREATE FINAL 26 x 26 MM LABEL
    # ========================================================

    label = Image.new(
        "RGB",
        (
            LABEL_SIZE,
            LABEL_SIZE
        ),
        "white"
    )


    # ========================================================
    # PLACE QR
    # ========================================================

    qr_x = (
        LABEL_SIZE
        - QR_SIZE
    ) // 2

    qr_y = 0

    label.paste(
        qr_image.convert("RGB"),
        (
            qr_x,
            qr_y
        )
    )


    # ========================================================
    # CREATE ASSET ID TEXT
    # ========================================================

    draw = ImageDraw.Draw(label)

    text_bbox = draw.textbbox(
        (0, 0),
        asset_id,
        font=font
    )

    original_text_width = (
        text_bbox[2]
        - text_bbox[0]
    )

    original_text_height = (
        text_bbox[3]
        - text_bbox[1]
    )


    # Temporary transparent text image
    temp_width = (
        original_text_width
        + mm_to_px(1)
    )

    temp_height = (
        original_text_height
        + mm_to_px(1)
    )

    text_image = Image.new(
        "RGBA",
        (
            temp_width,
            temp_height
        ),
        (255, 255, 255, 0)
    )

    text_draw = ImageDraw.Draw(
        text_image
    )


    # Draw Asset ID
    text_draw.text(
        (
            mm_to_px(0.5),
            -text_bbox[1]
            + mm_to_px(0.5)
        ),
        asset_id,
        fill="black",
        font=font
    )


    # ========================================================
    # STRETCH ASSET ID HORIZONTALLY
    # ========================================================

    target_text_width = (
        LABEL_SIZE
        - mm_to_px(0.5)
    )

    stretched_text = text_image.resize(
        (
            target_text_width,
            text_image.height
        ),
        Image.Resampling.LANCZOS
    )


    # ========================================================
    # CENTER ASSET ID
    # ========================================================

    text_x = (
        LABEL_SIZE
        - stretched_text.width
    ) // 2

    text_y = (
        QR_SIZE
        + TEXT_GAP
    )


    # ========================================================
    # PLACE ASSET ID
    # ========================================================

    label.paste(
        stretched_text,
        (
            text_x,
            text_y
        ),
        stretched_text
    )


    # ========================================================
    # SAVE LABEL
    # ========================================================

    label.save(
        output_file,
        format="PNG",
        dpi=(DPI, DPI)
    )


    # Add to current-run list
    new_qr_files.append(
        output_file
    )

    generated += 1


# ============================================================
# SAVE CURRENT-RUN QR LIST
# ============================================================

with open(
    NEW_QR_LIST,
    "w",
    encoding="utf-8"
) as file:

    for qr_file in new_qr_files:

        file.write(
            str(qr_file)
            + "\n"
        )


# ============================================================
# SUMMARY
# ============================================================

print()
print("==============================================")
print("QR GENERATION COMPLETE")
print("==============================================")
print(
    f"Generated / regenerated : {generated}"
)
print(
    f"Existing skipped        : {skipped}"
)
print("Complete label          : 26 x 26 mm")
print(
    f"QR size                 : "
    f"{QR_MM} x {QR_MM} mm"
)
print(
    f"Output folder           : {OUTPUT_DIR}"
)
print(
    f"New QR list             : {NEW_QR_LIST}"
)
print("==============================================")