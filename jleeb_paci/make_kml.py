"""Create jleeb_paci.kml from jleeb_paci.xlsx (rows that have LAT/LONG filled)."""
from xml.sax.saxutils import escape
from openpyxl import load_workbook

ws = load_workbook("jleeb_paci.xlsx").active
out = ['<?xml version="1.0" encoding="UTF-8"?>',
       '<kml xmlns="http://www.opengis.net/kml/2.2"><Document><name>Jleeb Al-Shuyoukh - Dangerous buildings (GM 2026/1727)</name>']
n = 0
for no, paci, lat, lon, area, block, plot, note in ws.iter_rows(min_row=2, values_only=True):
    if lat is None or lon is None: continue
    n += 1
    out.append(f"<Placemark><name>{escape(str(paci))}</name><description>PACI {paci} | Block {block} | Plot {escape(str(plot))}</description>"
               f"<Point><coordinates>{lon},{lat},0</coordinates></Point></Placemark>")
out.append("</Document></kml>")
open("jleeb_paci.kml", "w", encoding="utf-8").write("\n".join(out)); print(n, "placemarks")
