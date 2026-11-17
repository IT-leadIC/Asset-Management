from flask import (
    Flask,
    render_template_string,
    request,
    redirect,
    url_for,
    session,
    send_from_directory
)
from functools import wraps
from io import BytesIO
import base64
import os
import time

import requests
import msal
from openpyxl import load_workbook


app = Flask(__name__)


# ============================================================
# CONFIGURATION
# ============================================================

app.secret_key = os.environ.get(
    "FLASK_SECRET_KEY",
    "change-this-before-public-deployment"
)

ADMIN_USERNAME = os.environ.get(
    "ADMIN_USERNAME",
    "admin"
)

ADMIN_PASSWORD = os.environ.get(
    "ADMIN_PASSWORD",
    "AssetAdmin@2026"
)

AZURE_TENANT_ID = os.environ.get("AZURE_TENANT_ID")
AZURE_CLIENT_ID = os.environ.get("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET")
SHAREPOINT_FILE_URL = os.environ.get("SHAREPOINT_FILE_URL")


# ============================================================
# ASSET DATA CACHE
# ============================================================

# Keep Excel data in memory for 5 minutes
CACHE_DURATION_SECONDS = 300

_assets_cache = None
_assets_cache_time = 0


# ============================================================
# MICROSOFT GRAPH AUTHENTICATION
# ============================================================

def get_access_token():

    if not AZURE_TENANT_ID:
        raise ValueError("AZURE_TENANT_ID is missing.")

    if not AZURE_CLIENT_ID:
        raise ValueError("AZURE_CLIENT_ID is missing.")

    if not AZURE_CLIENT_SECRET:
        raise ValueError("AZURE_CLIENT_SECRET is missing.")

    authority = (
        f"https://login.microsoftonline.com/"
        f"{AZURE_TENANT_ID}"
    )

    client = msal.ConfidentialClientApplication(
        AZURE_CLIENT_ID,
        authority=authority,
        client_credential=AZURE_CLIENT_SECRET
    )

    result = client.acquire_token_for_client(
        scopes=[
            "https://graph.microsoft.com/.default"
        ]
    )

    if "access_token" not in result:

        error = result.get(
            "error_description",
            "Unable to obtain Microsoft Graph access token."
        )

        raise RuntimeError(error)

    return result["access_token"]


# ============================================================
# SHARE URL ENCODING
# ============================================================

def encode_share_url(url):

    encoded = base64.urlsafe_b64encode(
        url.encode("utf-8")
    ).decode("utf-8")

    encoded = encoded.rstrip("=")

    return "u!" + encoded


# ============================================================
# GET SHAREPOINT FILE INFORMATION
# ============================================================

def get_sharepoint_item():

    if not SHAREPOINT_FILE_URL:
        raise ValueError(
            "SHAREPOINT_FILE_URL is missing."
        )

    token = get_access_token()

    encoded_url = encode_share_url(
        SHAREPOINT_FILE_URL
    )

    url = (
        "https://graph.microsoft.com/v1.0/"
        f"shares/{encoded_url}/driveItem"
    )

    response = requests.get(
        url,
        headers={
            "Authorization": f"Bearer {token}"
        },
        timeout=30
    )

    if not response.ok:

        raise RuntimeError(
            "Unable to access SharePoint file. "
            f"HTTP {response.status_code}: "
            f"{response.text[:500]}"
        )

    return response.json()


# ============================================================
# DOWNLOAD EXCEL
# ============================================================

def download_excel():

    item = get_sharepoint_item()

    drive_id = (
        item.get("parentReference", {})
        .get("driveId")
    )

    item_id = item.get("id")

    # --------------------------------------------------------
    # Direct Graph drive item download
    # --------------------------------------------------------

    if drive_id and item_id:

        download_url = (
            "https://graph.microsoft.com/v1.0/"
            f"drives/{drive_id}/items/{item_id}/content"
        )

        token = get_access_token()

        response = requests.get(
            download_url,
            headers={
                "Authorization": f"Bearer {token}"
            },
            timeout=60
        )

        if response.ok:

            content = response.content

            if content.startswith(b"PK"):
                return content

    # --------------------------------------------------------
    # Fallback to Microsoft download URL
    # --------------------------------------------------------

    download_link = item.get(
        "@microsoft.graph.downloadUrl"
    )

    if not download_link:

        raise RuntimeError(
            "Unable to obtain a downloadable Excel file "
            "from SharePoint."
        )

    response = requests.get(
        download_link,
        timeout=60
    )

    if not response.ok:

        raise RuntimeError(
            "Unable to download Excel file. "
            f"HTTP {response.status_code}: "
            f"{response.text[:500]}"
        )

    content = response.content

    if not content.startswith(b"PK"):

        content_type = response.headers.get(
            "Content-Type",
            "unknown"
        )

        preview = content[:200].decode(
            "utf-8",
            errors="replace"
        )

        raise RuntimeError(
            "The SharePoint download did not return "
            "an Excel workbook.\n"
            f"Content-Type: {content_type}\n"
            f"Response preview: {preview}"
        )

    return content


# ============================================================
# FIND ASSET SHEET + HEADERS
# ============================================================

def find_asset_sheet(workbook):

    for ws in workbook.worksheets:

        for row_number in range(
            1,
            min(ws.max_row, 15) + 1
        ):

            headers = {}

            for cell in ws[row_number]:

                if cell.value is not None:

                    header = str(
                        cell.value
                    ).strip().lower()

                    headers[header] = cell.column

            if (
                "asset no" in headers
                and "employee" in headers
            ):

                return (
                    ws,
                    row_number,
                    headers
                )

    raise ValueError(
        "Could not find 'Asset No' and "
        "'Employee' columns in the workbook."
    )


# ============================================================
# READ ASSETS FROM EXCEL WITH CACHE
# ============================================================

def get_assets(force_refresh=False):

    global _assets_cache
    global _assets_cache_time

    now = time.time()

    # --------------------------------------------------------
    # USE CACHE WHEN VALID
    # --------------------------------------------------------

    if (
        not force_refresh
        and _assets_cache is not None
        and (now - _assets_cache_time)
        < CACHE_DURATION_SECONDS
    ):

        return _assets_cache

    # --------------------------------------------------------
    # LOAD FRESH DATA
    # --------------------------------------------------------

    print(
        "Loading latest Excel data from SharePoint..."
    )

    content = download_excel()

    workbook = load_workbook(
        BytesIO(content),
        data_only=True
    )

    ws, header_row, headers = find_asset_sheet(
        workbook
    )

    assets = []

    for row_number in range(
        header_row + 1,
        ws.max_row + 1
    ):

        asset_no = ws.cell(
            row=row_number,
            column=headers["asset no"]
        ).value

        if asset_no is None:
            continue

        asset_no = str(asset_no).strip()

        if not asset_no:
            continue

        def value(header):

            column = headers.get(
                header.lower()
            )

            if column is None:
                return ""

            cell_value = ws.cell(
                row=row_number,
                column=column
            ).value

            if cell_value is None:
                return ""

            return str(cell_value).strip()

        assets.append({
            "asset_id": asset_no,
            "asset_group": value("Asset Group"),
            "asset_type": value("Asset Type"),
            "description": value("Asset Description"),
            "location": value("Location"),
            "status": value("Status"),
            "brand": value("Brand"),
            "model": value("Model"),
            "serial_no": value("Serial No"),
            "purchase_date": value("Purchase Date"),
            "invoice_no": value("Invoice No"),
            "original_value": value("Original Value"),
            "value": value("Value"),
            "warranty_date": value("Warranty Date"),
            "warranty_status": value("Warranty Status"),
            "employee": value("Employee"),
            "remarks": value("Remarks")
        })

    workbook.close()

    assets.sort(
        key=lambda x: x["asset_id"].lower()
    )

    # --------------------------------------------------------
    # STORE IN CACHE
    # --------------------------------------------------------

    _assets_cache = assets
    _assets_cache_time = time.time()

    print(
        f"Excel loaded and cached: "
        f"{len(assets)} assets"
    )

    return assets


# ============================================================
# EMPLOYEE LIST
# ============================================================

def get_employee_list(force_refresh=False):

    assets = get_assets(
        force_refresh=force_refresh
    )

    employees = set()

    for asset in assets:

        employee = asset.get(
            "employee",
            ""
        ).strip()

        if employee:
            employees.add(employee)

    return sorted(
        employees,
        key=str.lower
    )


# ============================================================
# UPDATE EMPLOYEE IN EXCEL
# ============================================================

def update_excel_employee(
    asset_id,
    employee
):

    global _assets_cache
    global _assets_cache_time

    # --------------------------------------------------------
    # GET LATEST FILE INFORMATION
    # --------------------------------------------------------

    item = get_sharepoint_item()

    drive_id = (
        item.get("parentReference", {})
        .get("driveId")
    )

    item_id = item.get("id")

    if not drive_id or not item_id:

        raise RuntimeError(
            "Could not determine SharePoint drive "
            "or file ID."
        )

    # --------------------------------------------------------
    # DOWNLOAD LATEST WORKBOOK
    # --------------------------------------------------------

    content = download_excel()

    workbook = load_workbook(
        BytesIO(content),
        data_only=False
    )

    ws, header_row, headers = find_asset_sheet(
        workbook
    )

    asset_column = headers["asset no"]
    employee_column = headers["employee"]

    found = False

    # --------------------------------------------------------
    # FIND ASSET
    # --------------------------------------------------------

    for row_number in range(
        header_row + 1,
        ws.max_row + 1
    ):

        current_asset = ws.cell(
            row=row_number,
            column=asset_column
        ).value

        if current_asset is None:
            continue

        current_asset = str(
            current_asset
        ).strip()

        if current_asset.lower() == str(
            asset_id
        ).strip().lower():

            ws.cell(
                row=row_number,
                column=employee_column
            ).value = employee

            found = True
            break

    if not found:

        workbook.close()

        raise ValueError(
            f"Asset '{asset_id}' was not found "
            "in the Excel workbook."
        )

    # --------------------------------------------------------
    # RECALCULATE EXCEL FORMULAS
    # --------------------------------------------------------

    try:

        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"

    except Exception:
        pass

    # --------------------------------------------------------
    # SAVE TO MEMORY
    # --------------------------------------------------------

    output = BytesIO()

    workbook.save(output)
    workbook.close()

    file_bytes = output.getvalue()

    # --------------------------------------------------------
    # UPLOAD TO SHAREPOINT
    # --------------------------------------------------------

    token = get_access_token()

    upload_url = (
        "https://graph.microsoft.com/v1.0/"
        f"drives/{drive_id}/items/{item_id}/content"
    )

    request_headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": (
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        )
    }

    etag = item.get("@odata.etag")

    if etag:
        request_headers["If-Match"] = etag

    response = requests.put(
        upload_url,
        headers=request_headers,
        data=file_bytes,
        timeout=60
    )

    if response.status_code == 412:

        raise RuntimeError(
            "The Excel file was changed by someone else "
            "while this update was being processed. "
            "Please refresh and try again."
        )

    if not response.ok:

        raise RuntimeError(
            "Unable to save changes to SharePoint. "
            f"HTTP {response.status_code}: "
            f"{response.text[:500]}"
        )

    # --------------------------------------------------------
    # CLEAR CACHE
    # --------------------------------------------------------

    _assets_cache = None
    _assets_cache_time = 0

    print(
        "Employee updated successfully. "
        "Asset cache cleared."
    )


# ============================================================
# LOGIN REQUIRED
# ============================================================

def login_required(f):

    @wraps(f)
    def decorated_function(
        *args,
        **kwargs
    ):

        if not session.get(
            "admin_logged_in"
        ):

            return redirect(
                url_for("login")
            )

        return f(*args, **kwargs)

    return decorated_function


# ============================================================
# HOME
# ============================================================

@app.route("/")
def home():

    return redirect(
        url_for("login")
    )


# ============================================================
# SERVE LEADIC LOGO
# ============================================================

@app.route("/logo.png")
def logo_file():
    return send_from_directory(
        os.path.dirname(os.path.abspath(__file__)),
        "logo.png"
    )


# ============================================================
# HEALTH CHECK
# ============================================================

@app.route("/health")
def health():
    return "OK", 200


# ============================================================
# PUBLIC ASSET DETAILS
# QR SCAN ROUTE
# ============================================================

@app.route("/asset/<asset_id>")
def asset_details(asset_id):

    try:

        assets = get_assets()

        asset = next(
            (
                item
                for item in assets
                if item["asset_id"].strip().lower()
                == asset_id.strip().lower()
            ),
            None
        )

        if asset is None:

            return render_template_string(
                ASSET_NOT_FOUND_TEMPLATE,
                asset_id=asset_id
            ), 404

        return render_template_string(
            ASSET_TEMPLATE,
            asset=asset
        )

    except Exception as e:

        return render_template_string(
            ERROR_TEMPLATE,
            message=str(e)
        ), 500


# ============================================================
# LOGIN
# ============================================================

@app.route(
    "/login",
    methods=["GET", "POST"]
)
def login():

    if request.method == "POST":

        username = request.form.get(
            "username",
            ""
        ).strip()

        password = request.form.get(
            "password",
            ""
        )

        if (
            username == ADMIN_USERNAME
            and password == ADMIN_PASSWORD
        ):

            session[
                "admin_logged_in"
            ] = True

            return redirect(
                url_for("admin")
            )

        return render_template_string(
            LOGIN_TEMPLATE,
            error=True
        )

    return render_template_string(
        LOGIN_TEMPLATE,
        error=False
    )


# ============================================================
# LOGOUT
# ============================================================

@app.route("/logout")
def logout():

    session.clear()

    return redirect(
        url_for("login")
    )


# ============================================================
# ADMIN DASHBOARD
# ============================================================

@app.route("/admin")
@login_required
def admin():

    try:

        force_refresh = (
            request.args.get("refresh") == "1"
        )

        assets = get_assets(
            force_refresh=force_refresh
        )

        return render_template_string(
            DASHBOARD_TEMPLATE,
            assets=assets
        )

    except Exception as e:

        return render_template_string(
            ERROR_TEMPLATE,
            message=str(e)
        ), 500


# ============================================================
# EDIT EMPLOYEE
# ============================================================

@app.route(
    "/admin/edit/<asset_id>",
    methods=["GET", "POST"]
)
@login_required
def edit_asset(asset_id):

    try:

        assets = get_assets()

        asset = next(
            (
                item
                for item in assets
                if item["asset_id"].strip().lower()
                == asset_id.strip().lower()
            ),
            None
        )

        if asset is None:

            return "Asset not found", 404

        employees = get_employee_list()

        if request.method == "POST":

            employee = request.form.get(
                "employee",
                ""
            ).strip()

            if (
                employee
                and employee not in employees
            ):

                return (
                    "Invalid employee selection.",
                    400
                )

            update_excel_employee(
                asset_id,
                employee
            )

            return redirect(
                url_for("admin")
            )

        return render_template_string(
            EDIT_TEMPLATE,
            asset=asset,
            employees=employees
        )

    except Exception as e:

        return render_template_string(
            ERROR_TEMPLATE,
            message=str(e)
        ), 500


# ============================================================
# LOGIN TEMPLATE
# ============================================================

LOGIN_TEMPLATE = """
<!DOCTYPE html>
<html>

<head>

<meta name="viewport"
      content="width=device-width, initial-scale=1">

<title>Asset Management Login</title>

<style>

body {
    font-family: Arial, sans-serif;
    background: #f3f4f6;
    padding: 20px;
}

.card {
    max-width: 400px;
    margin: 70px auto;
    background: white;
    padding: 25px;
    border-radius: 15px;
    box-shadow: 0 3px 12px rgba(0,0,0,0.15);
}

h2 {
    text-align: center;
}

label {
    display: block;
    margin-top: 15px;
    font-weight: bold;
}

input {
    width: 100%;
    padding: 10px;
    margin-top: 5px;
    box-sizing: border-box;
}

button {
    width: 100%;
    padding: 12px;
    margin-top: 20px;
    cursor: pointer;
}

.error {
    color: red;
    text-align: center;
}

</style>

</head>

<body>

<div class="card">

<h2>ADMIN LOGIN</h2>

{% if error %}

<p class="error">
Invalid username or password
</p>

{% endif %}

<form method="POST">

<label>Username</label>

<input
    type="text"
    name="username"
    required
>

<label>Password</label>

<input
    type="password"
    name="password"
    required
>

<button type="submit">
LOGIN
</button>

</form>

</div>

</body>

</html>
"""


# ============================================================
# DASHBOARD TEMPLATE
# ============================================================

DASHBOARD_TEMPLATE = """
<!DOCTYPE html>
<html>

<head>

<meta name="viewport"
      content="width=device-width, initial-scale=1">

<title>Asset Management</title>

<style>

body {
    font-family: Arial, sans-serif;
    margin: 20px;
    background: #f5f5f5;
}

h1 {
    text-align: center;
}

.top {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 15px;
    gap: 10px;
}

.refresh,
.logout {
    padding: 8px 12px;
    background: #ddd;
    border-radius: 5px;
    text-decoration: none;
    color: black;
}

.table-container {
    overflow-x: auto;
}

table {
    width: 100%;
    border-collapse: collapse;
    background: white;
    min-width: 1000px;
}

th,
td {
    padding: 10px;
    border: 1px solid #ddd;
    text-align: left;
    white-space: nowrap;
}

th {
    background: #eee;
}

a {
    text-decoration: none;
}

</style>

</head>

<body>

<h1>Asset Management</h1>

<div class="top">

<div>
    <b>Total Assets:</b>
    {{ assets|length }}
</div>

<div>

<a class="refresh"
   href="{{ url_for('admin', refresh='1') }}">
   Refresh
</a>

<a class="logout"
   href="{{ url_for('logout') }}">
   Logout
</a>

</div>

</div>

<div class="table-container">

<table>

<tr>

<th>Asset No</th>
<th>Asset Type</th>
<th>Brand</th>
<th>Model</th>
<th>Serial No</th>
<th>Status</th>
<th>Employee</th>
<th>Warranty Status</th>
<th>Remarks</th>
<th>Action</th>

</tr>

{% for asset in assets %}

<tr>

<td>{{ asset["asset_id"] }}</td>

<td>{{ asset["asset_type"] }}</td>

<td>{{ asset["brand"] }}</td>

<td>{{ asset["model"] }}</td>

<td>{{ asset["serial_no"] }}</td>

<td>{{ asset["status"] }}</td>

<td>{{ asset["employee"] }}</td>

<td>{{ asset["warranty_status"] }}</td>

<td>{{ asset["remarks"] }}</td>

<td>

<a href="{{ url_for(
    'edit_asset',
    asset_id=asset['asset_id']
) }}">
Edit Employee
</a>

</td>

</tr>

{% endfor %}

</table>

</div>

</body>

</html>
"""


# ============================================================
# EDIT TEMPLATE
# ============================================================

EDIT_TEMPLATE = """
<!DOCTYPE html>
<html>

<head>

<meta name="viewport"
      content="width=device-width, initial-scale=1">

<title>Edit Employee</title>

<style>

body {
    font-family: Arial, sans-serif;
    background: #f5f5f5;
    padding: 20px;
}

.card {
    max-width: 500px;
    margin: auto;
    background: white;
    padding: 25px;
    border-radius: 15px;
}

label {
    display: block;
    margin-top: 15px;
    font-weight: bold;
}

select {
    width: 100%;
    padding: 10px;
    margin-top: 5px;
    box-sizing: border-box;
}

button {
    margin-top: 20px;
    width: 100%;
    padding: 12px;
    cursor: pointer;
}

.back {
    display: inline-block;
    margin-top: 15px;
}

</style>

</head>

<body>

<div class="card">

<h2>Edit Employee</h2>

<p>
<b>Asset No:</b>
{{ asset["asset_id"] }}
</p>

<p>
<b>Asset Type:</b>
{{ asset["asset_type"] }}
</p>

<p>
<b>Model:</b>
{{ asset["model"] }}
</p>

<p>
<b>Serial No:</b>
{{ asset["serial_no"] }}
</p>

<form method="POST">

<label>Employee</label>

<select name="employee">

<option value="">
-- Unassigned --
</option>

{% for employee in employees %}

<option
    value="{{ employee }}"
    {% if employee == asset["employee"] %}
    selected
    {% endif %}
>
{{ employee }}
</option>

{% endfor %}

</select>

<button type="submit">
SAVE CHANGES
</button>

</form>

<a class="back"
   href="{{ url_for('admin') }}">
← Back to Dashboard
</a>

</div>

</body>

</html>
"""


# ============================================================
# PUBLIC ASSET DETAILS TEMPLATE
# QR SCAN PAGE
# ============================================================

ASSET_TEMPLATE = """
<!DOCTYPE html>
<html>

<head>

<meta name="viewport"
      content="width=device-width, initial-scale=1">

<title>Asset Details</title>

<style>

body {
    margin: 0;
    padding: 20px;
    background: #f4f6f8;
    font-family: Arial, sans-serif;
}

.card {
    max-width: 520px;
    margin: 20px auto;
    background: #ffffff;
    border-radius: 16px;
    padding: 24px;
    box-shadow: 0 4px 16px rgba(0,0,0,0.12);
}


/* ============================================================
   HEADER
============================================================ */

.header {
    margin-bottom: 20px;
}

.logo-row {
    height: 60px;
    display: flex;
    align-items: flex-start;
    justify-content: flex-start;
}

.logo {
    width: 120px;
    max-height: 55px;
    height: auto;
    object-fit: contain;
}

.header-title {
    text-align: center;
    font-size: 24px;
    font-weight: 700;
    color: #222;
    margin-top: 5px;
}


/* ============================================================
   ASSET ID
============================================================ */

.asset-id {
    text-align: center;
    font-size: 24px;
    font-weight: 700;
    color: #111;
    padding: 12px 0 18px 0;
    border-bottom: 2px solid #eeeeee;
    margin-bottom: 4px;
}


/* ============================================================
   DETAILS TABLE
============================================================ */

.detail-row {
    display: grid;
    grid-template-columns: 42% 58%;
    align-items: center;
    min-height: 48px;
    border-bottom: 1px solid #e8e8e8;
}

.detail-label {
    font-weight: 600;
    color: #555;
    text-align: left;
    padding-right: 12px;
}

.detail-value {
    color: #111;
    text-align: right;
    font-weight: 500;
    word-break: break-word;
}


/* ============================================================
   MOBILE
============================================================ */

@media (max-width: 480px) {

    body {
        padding: 12px;
    }

    .card {
        margin: 10px auto;
        padding: 18px;
    }

    .logo {
        width: 105px;
        max-height: 50px;
    }

    .header-title {
        font-size: 22px;
    }

    .asset-id {
        font-size: 22px;
    }

    .detail-row {
        grid-template-columns: 40% 60%;
        min-height: 46px;
    }

    .detail-label,
    .detail-value {
        font-size: 14px;
    }

}

</style>

</head>


<body>

<div class="card">


<!-- ============================================================
     LOGO + TITLE
============================================================ -->

<div class="header">

    <div class="logo-row">

        <img
            class="logo"
            src="{{ url_for('logo_file') }}"
            alt="LeadIC Design"
        >

    </div>

    <div class="header-title">
        Asset Details
    </div>

</div>


<!-- ============================================================
     ASSET ID
============================================================ -->

<div class="asset-id">
    {{ asset["asset_id"] }}
</div>


<!-- ============================================================
     DETAILS
============================================================ -->

<div class="detail-row">

    <div class="detail-label">
        Asset Type
    </div>

    <div class="detail-value">
        {{ asset["asset_type"] }}
    </div>

</div>


<div class="detail-row">

    <div class="detail-label">
        Model
    </div>

    <div class="detail-value">
        {{ asset["model"] }}
    </div>

</div>


<div class="detail-row">

    <div class="detail-label">
        Serial No
    </div>

    <div class="detail-value">
        {{ asset["serial_no"] }}
    </div>

</div>


<div class="detail-row">

    <div class="detail-label">
        Status
    </div>

    <div class="detail-value">
        {{ asset["status"] }}
    </div>

</div>


<div class="detail-row">

    <div class="detail-label">
        Employee
    </div>

    <div class="detail-value">
        {{ asset["employee"] }}
    </div>

</div>


</div>

</body>

</html>
"""


# ============================================================
# ASSET NOT FOUND TEMPLATE
# ============================================================

ASSET_NOT_FOUND_TEMPLATE = """
<!DOCTYPE html>
<html>

<head>

<meta name="viewport"
      content="width=device-width, initial-scale=1">

<title>Asset Not Found</title>

<style>

body {
    font-family: Arial, sans-serif;
    background: #f5f5f5;
    padding: 20px;
}

.card {
    max-width: 500px;
    margin: 60px auto;
    background: white;
    padding: 30px;
    border-radius: 15px;
    text-align: center;
}

h2 {
    color: #b00020;
}

.asset {
    font-size: 20px;
    font-weight: bold;
}

</style>

</head>

<body>

<div class="card">

<h2>Asset Not Found</h2>

<p>
No asset was found for:
</p>

<div class="asset">
{{ asset_id }}
</div>

</div>

</body>

</html>
"""


# ============================================================
# ERROR TEMPLATE
# ============================================================

ERROR_TEMPLATE = """
<!DOCTYPE html>
<html>

<head>

<meta name="viewport"
      content="width=device-width, initial-scale=1">

<title>Asset Management Error</title>

<style>

body {
    font-family: Arial, sans-serif;
    background: #f5f5f5;
    padding: 20px;
}

.card {
    max-width: 700px;
    margin: 60px auto;
    background: white;
    padding: 25px;
    border-radius: 15px;
}

.error {
    color: #b00020;
    white-space: pre-wrap;
}

</style>

</head>

<body>

<div class="card">

<h2>Asset Management Error</h2>

<p class="error">
{{ message }}
</p>

<a href="{{ url_for('admin') }}">
Back to Dashboard
</a>

</div>

</body>

</html>
"""


# ============================================================
# START SERVER
# ============================================================

if __name__ == "__main__":

    port = int(
        os.environ.get(
            "PORT",
            5000
        )
    )

    app.run(
        host="0.0.0.0",
        port=port
    )
