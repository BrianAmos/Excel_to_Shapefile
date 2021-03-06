# Excel_to_Shapefile
Convert .xlsx Excel spreadsheets to Esri shapefiles.

By Brian Amos, twitter @brianamos.

Built in Python 3.7, requires the openpyxl and fiona packages.

The first sheet in the Excel workbook should be your "map," where each cell is labeled by the ID of the polygon that it will be a part of in the final product.

The optional second sheet in the workbook is the attribute table. Row one is the field names, the first of which must be "id", which will pair with the IDs used in the first sheet. Row two are the field types - check fiona.FIELD_TYPES_MAP for options, but str, int, and float are common choices. The remaining rows are the data, and each ID used in the first sheet must have a row in the second.

See the included .xlsx for an example.
