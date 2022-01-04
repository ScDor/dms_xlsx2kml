# dms_xlsx2kml
Converts a XLSX including DMS coordinates to a KML file


# How to use
1. Create an excel file where the first six columns represent lat_D, lat_M, lat_S, lng_D, lng_M, lng_S 
For example, an excel with one point of the chords bridge in Jerusalem (31°47'20.0"N 35°11'59.9"E) should be
 
| d1 | m1 | s1   | d2 | m2 | s2   |
|----|----|------|----|----|------|
| 31 | 47 | 20.0 | 35 | 11 | 59.9 |


2. Modify the `.py` file main to run on any excel files you may have.
