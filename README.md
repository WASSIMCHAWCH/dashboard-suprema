# Suprema Door Access Dashboard

This dashboard reads the shared `doors` folder and builds a live access matrix in the browser.

## Run

```
python app.py
```

Override the path if needed:


```

Open `http://127.0.0.1:8000` in your browser.
Open `http://127.0.0.1:8000/comper` to compare current access with the support JSON.

Use the **Export Excel** button to download a formatted `.xlsx` with green (access) and red (no access) cells.
The exported Excel sheet is protected with password `0000`.

Departments are loaded from the `CDC` sheet in the support file by default. Override with:



The comparison reads the `CDC` sheet and merges duplicate door columns (same L-code) so a door is granted if any duplicate column has an X.
The JSON report is saved to `exports/doors_reference.json` inside the project automatically when `/comper` loads.
The `/comper` page reads that JSON file by default. Override with:

```
$env:SUPPORT_JSON_PATH = "C:\\path\\to\\doors_reference.json"
python app.py
```

Relative paths for `SUPPORT_JSON_PATH`, `DOORS_PATH`, and `SUPREMA_PATH` are resolved from the project folder.

## Optional department mapping

If you want a Department column, provide a CSV with headers `function,department`:

```
function,department
Technicien AQ,AQ
Operateur de conditionnement,PROD
```

Set:

```
$env:DEPARTMENTS_FILE = "C:\\path\\to\\departments.csv"
```

If `departments.csv` exists inside the `doors` folder, it will be used automatically when `SUPREMA_PATH` is not found.

## Optional door metadata (layout + headers)

To control door labels, groups, codes, and ordering in the export, add `door_metadata.csv` to the `doors` folder (or set `DOOR_METADATA_FILE`).

Example:

```
file,label,group,code,order
Administra IN.csv,Admin-In,Administration,L01,1
Administra OUT.csv,Admin-Out,Administration,L02,2
MAG IN.csv,MAG In,MAG,L07,3
```
