import simplekml
from openpyxl import load_workbook


def _dms2dd(d: float, m: float, s: float) -> float:
    """
    Converts a DMS coordinate to DD
    :param d: degrees
    :param m: minutes
    :param s: seconds
    :return: equivalent in decimal degrees
    """
    return d + m / 60 + s / 3600


def xlsx2kml(xlsx_path: str, kml_path: str) -> None:
    """
    :param xlsx_path: path to an xlsx file, with a header line (ignored) and lines of the format
            lat_d, lat_m, lat_s, lng_d, lng_m, lng_s

            For example, an excel with one point of the chords bridge in Jerusalem (31°47'20.0"N 35°11'59.9"E)
             will look like:
            | ----- | ----- | ----- | ----- | ----- | ------ |
            |   d1	|   m1  |   s1  |   d2  |	m2	|   s2   |  # this is the head column
            |   31  |	47  |	20  |	35  |	11  |	59.9 |
            | ----- | ----- | ----- | ----- | ----- | ------ |

    :param kml_path: path to a kml file (overwritten if exists)
    """
    wb = load_workbook(filename=xlsx_path)
    sheet = wb.active  # uses the active sheet

    kml = simplekml.Kml()

    for i, row in enumerate(sheet.rows):
        if i == 0 or row[0].value is None:
            continue  # ignores header row and empty rows

        lat = _dms2dd(row[0].value, row[1].value, row[2].value)
        lng = _dms2dd(row[3].value, row[4].value, row[5].value)

        kml.newpoint(name=i, coords=[(lat, lng)])

    kml.save(kml_path)


if __name__ == '__main__':
    xlsx2kml('example.xlsx', 'chords_bridge.kml')
