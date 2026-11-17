from pathlib import Path
from math import ceil

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader


# ============================================================
# PROJECT SETTINGS
# ============================================================

PROJECT_DIR = Path(__file__).resolve().parent

QR_FOLDER = PROJECT_DIR / "QR_Codes"

NEW_QR_LIST = PROJECT_DIR / "new_qr_files.txt"

OUTPUT_FOLDER = PROJECT_DIR / "Labels"

OUTPUT_FOLDER.mkdir(
    exist_ok=True
)


# ============================================================
# LABEL GENERATION MODE
# ============================================================

# True  = Print only newly generated QR labels
# False = Print ALL QR labels
PRINT_NEW_ONLY = True


# ============================================================
# PDF OUTPUT NAME
# ============================================================

if PRINT_NEW_ONLY:

    OUTPUT_PDF = (
        OUTPUT_FOLDER
        / "New_A4_QR_Labels.pdf"
    )

else:

    OUTPUT_PDF = (
        OUTPUT_FOLDER
        / "All_A4_QR_Labels.pdf"
    )


# ============================================================
# A4 PAGE SIZE
# ============================================================

PAGE_WIDTH, PAGE_HEIGHT = A4


# ============================================================
# LABEL SIZE
# ============================================================

LABEL_WIDTH = 26 * mm
LABEL_HEIGHT = 26 * mm


# ============================================================
# GAP BETWEEN LABELS
# ============================================================

HORIZONTAL_GAP = 3 * mm
VERTICAL_GAP = 3 * mm


# ============================================================
# A4 GRID
# ============================================================

COLUMNS = 7
ROWS = 10

LABELS_PER_PAGE = (
    COLUMNS * ROWS
)


# ============================================================
# GET QR FILES
# ============================================================

if PRINT_NEW_ONLY:

    # --------------------------------------------------------
    # NEW LABELS ONLY
    # --------------------------------------------------------

    if not NEW_QR_LIST.exists():

        raise FileNotFoundError(
            "new_qr_files.txt was not found.\n"
            "Run generate_qr.py first."
        )

    with open(
        NEW_QR_LIST,
        "r",
        encoding="utf-8"
    ) as file:

        qr_files = [
            Path(line.strip())
            for line in file
            if line.strip()
        ]

else:

    # --------------------------------------------------------
    # ALL LABELS
    # --------------------------------------------------------

    if not QR_FOLDER.exists():

        raise FileNotFoundError(
            f"QR folder not found:\n{QR_FOLDER}"
        )

    qr_files = sorted(
        [
            file
            for file in QR_FOLDER.iterdir()
            if file.is_file()
            and file.suffix.lower() == ".png"
        ],
        key=lambda x: x.stem.lower()
    )


# ============================================================
# VERIFY QR FILES
# ============================================================

valid_qr_files = []

for qr_file in qr_files:

    if qr_file.exists():

        valid_qr_files.append(
            qr_file
        )

    else:

        print(
            f"Warning: QR file not found: "
            f"{qr_file}"
        )


qr_files = valid_qr_files


if not qr_files:

    print()
    print(
        "No QR labels available for printing."
    )
    print(
        "There is nothing new to print."
    )
    raise SystemExit


# ============================================================
# CALCULATE PAGE COUNT
# ============================================================

total_pages = ceil(
    len(qr_files)
    / LABELS_PER_PAGE
)


# ============================================================
# CREATE PDF
# ============================================================

pdf = canvas.Canvas(
    str(OUTPUT_PDF),
    pagesize=A4
)

pdf.setTitle(
    "IT Asset QR Labels"
)


# ============================================================
# CALCULATE GRID SIZE
# ============================================================

grid_width = (
    COLUMNS * LABEL_WIDTH
    + (COLUMNS - 1)
    * HORIZONTAL_GAP
)

grid_height = (
    ROWS * LABEL_HEIGHT
    + (ROWS - 1)
    * VERTICAL_GAP
)


# ============================================================
# CENTER GRID ON A4
# ============================================================

start_x = (
    PAGE_WIDTH
    - grid_width
) / 2

start_y = (
    PAGE_HEIGHT
    + grid_height
) / 2


# ============================================================
# CREATE PAGES
# ============================================================

for page_number in range(
    total_pages
):

    start_index = (
        page_number
        * LABELS_PER_PAGE
    )

    end_index = min(
        start_index
        + LABELS_PER_PAGE,
        len(qr_files)
    )

    page_files = qr_files[
        start_index:end_index
    ]


    # --------------------------------------------------------
    # PLACE LABELS
    # --------------------------------------------------------

    for index, qr_file in enumerate(
        page_files
    ):

        column = (
            index
            % COLUMNS
        )

        row = (
            index
            // COLUMNS
        )


        # X position
        x = (
            start_x
            + column
            * (
                LABEL_WIDTH
                + HORIZONTAL_GAP
            )
        )


        # Y position
        y = (
            start_y
            - LABEL_HEIGHT
            - row
            * (
                LABEL_HEIGHT
                + VERTICAL_GAP
            )
        )


        # ----------------------------------------------------
        # PLACE QR LABEL
        # EXACT 26 x 26 MM
        # ----------------------------------------------------

        image = ImageReader(
            str(qr_file)
        )

        pdf.drawImage(
            image,
            x,
            y,
            width=LABEL_WIDTH,
            height=LABEL_HEIGHT,
            preserveAspectRatio=False,
            mask="auto"
        )


        # ----------------------------------------------------
        # CUT BORDER
        # ----------------------------------------------------

        pdf.setLineWidth(
            0.25
        )

        pdf.setDash(
            1,
            1
        )

        pdf.rect(
            x,
            y,
            LABEL_WIDTH,
            LABEL_HEIGHT,
            stroke=1,
            fill=0
        )

        pdf.setDash()


    # --------------------------------------------------------
    # PAGE NUMBER
    # --------------------------------------------------------

    pdf.setFont(
        "Helvetica",
        7
    )

    pdf.drawCentredString(
        PAGE_WIDTH / 2,
        2.5 * mm,
        f"Page {page_number + 1} "
        f"of {total_pages}"
    )


    # --------------------------------------------------------
    # NEXT PAGE
    # --------------------------------------------------------

    pdf.showPage()


# ============================================================
# SAVE PDF
# ============================================================

pdf.save()


# ============================================================
# COMPLETE
# ============================================================

print()
print("==============================================")
print("A4 QR LABEL PDF CREATED")
print("==============================================")
print(
    f"Labels       : {len(qr_files)}"
)
print(
    f"Pages        : {total_pages}"
)
print(
    "Labels/page  : 70"
)
print(
    "Grid         : 7 columns x 10 rows"
)
print(
    "Label size   : 26 x 26 mm"
)
print(
    "Gap          : 3 mm"
)
print(
    f"PDF          : {OUTPUT_PDF}"
)
print("==============================================")