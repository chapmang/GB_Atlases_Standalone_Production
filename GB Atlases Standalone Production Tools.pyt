import arcpy


class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "PDFExport"
        self.alias = "PDF Export"

        # List of tool classes associated with this toolbox
        self.tools = [BatchPDF]


class BatchPDF(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "AA Batch PDF Export"
        self.description = "Batch Export all PDFs from a selected series"
        self.canRunInBackground = False
        # self.stylesheet = "AADlgContent.xsl"


    def getParameterInfo(self):

        product_definitions_table = r"Database Connections\CartoLive_GreatBritain.sde\PRODUCT_DEFINITIONS"

        """Define parameter definitions"""
        # Parameter 0
        product_list = arcpy.Parameter(
            displayName="Product",
            name="product",
            datatype="GPString",
            parameterType="Required",
            direction="input")
        product_list.filter.type = "ValueList"

        # Get the fields from the input
        fields = arcpy.ListFields(product_definitions_table)
        # Create a fieldinfo object
        field_info = arcpy.FieldInfo()
        # Iterate through the fields and set them to field_info
        for field in fields:
            if field.name == "PRODUCT_ID":
                field_info.addField(field.name, field.name, "VISIBLE", "")
            elif field.name == "NAME":
                field_info.addField(field.name, field.name, "VISIBLE", "")
            elif field.name == "SCALE_SIZE":
                field_info.addField(field.name, field.name, "HIDDEN", "")
            elif field.name == "SHORT_NAME":
                field_info.addField(field.name, field.name, "HIDDEN", "")
            elif field.name == "PRODUCT_TYPE":
                field_info.addField(field.name, field.name, "HIDDEN", "")
            elif field.name == "created_user":
                field_info.addField(field.name, field.name, "HIDDEN", "")
            elif field.name == "created_date":
                field_info.addField(field.name, field.name, "HIDDEN", "")
            elif field.name == "last_edited_user":
                field_info.addField(field.name, field.name, "HIDDEN", "")
            elif field.name == "last_edited_date":
                field_info.addField(field.name, field.name, "HIDDEN", "")

        # Get list of Book Products
        if arcpy.Exists("book_products_view"):
            arcpy.AddMessage('hello')
        else:
           arcpy.MakeTableView_management(product_definitions_table, "book_products_view", "PRODUCT_TYPE = 'BOOK'", "", field_info)

        product_list_values = []
        with arcpy.da.SearchCursor("book_products_view", "NAME") as cur:
            for row in cur:
                product_list_values.append(row[0])
                arcpy.AddMessage(row[0])
        product_list.filter.list = product_list_values

        # Parameter 1
        pagination_file = arcpy.Parameter(
            displayName="Pagination File",
            name="pagination_file",
            datatype="DEFeatureClass",
            parameterType="Required",
            direction="input"
        )

        # Parameter 2
        template_file = arcpy.Parameter(
            displayName="Choose the template MXD",
            name="template_file",
            datatype="DEMapDocument",
            parameterType="Required",
            direction="input")

        # Parameter 3
        settings_file = arcpy.Parameter(
            displayName="Choose the production settings XML file",
            name="settings_file",
            datatype="DEFile",
            parameterType="Required",
            direction="input")

        # Parameter 4
        all_pages = arcpy.Parameter(
            displayName="All Pages",
            name="all_pages",
            datatype="GPBoolean",
            parameterType="Required",
            direction="input",
        )
        all_pages.value = "True"

        # Parameter 5
        page_range = arcpy.Parameter(
            displayName="Page Range",
            name="page_range",
            datatype="GPString",
            parameterType="Optional",
            direction="input",
            enabled=False)

        # Parameter 6
        destination_folder = arcpy.Parameter(
            displayName="Choose the destination directory for the PDFs",
            name="destination_folder",
            datatype="DEFolder",
            parameterType="Required",
            direction="input")

        return[product_list,
               template_file,
               pagination_file,
               settings_file,
               all_pages,
               page_range,
               destination_folder]

    arcpy.Delete_management("book_products_view")
    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        try:
            if arcpy.CheckExtension("foundation") == "Available":
                # Check out Production Mapping license
                arcpy.CheckOutExtension("foundation")
            else:
                raise Exception
        except:
            return False
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""

        if parameters[4].value:
            parameters[5].enabled = False
        else:
            parameters[5].enabled = True

        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
        # -------------------------------------------------------------------------------
        # Name:         Batch PDF Export
        # Purpose:      Traverses the Product Library, checks out the required product
        #               and exports the file to a Production PDF using AA default
        #               settings. The required product is then checked back in to the
        #               Product Library, with an os based clean to catch any copys not
        #               removed during check in
        #
        # Author:      chapmang
        #
        # Created:     15/03/2019
        # Copyright:   (c) chapmang 2019
        # Licence:
        # -------------------------------------------------------------------------------

        import arcpy
        import arcpyproduction
        import os
        import errno
        import re
        import sys
        import gc
        arcpy.overwriteOutput = True

        # Fetch values from user input
        product_name = parameters[0].valueAsText
        template_file = parameters[1].valueAsText
        pagination_file = parameters[2].valueAsText
        settings_file = parameters[3].valueAsText
        all_pages = parameters[4].valueAsText
        page_range = parameters[5].valueAsText
        opath = parameters[6].valueAsText

        # Fetch the product ID
        product_definitions_table = r"Database Connections\CartoLive_GreatBritain.sde\PRODUCT_DEFINITIONS"
        arcpy.MakeTableView_management(product_definitions_table, "book_products")
        with arcpy.da.SearchCursor("book_products", "PRODUCT_ID", "NAME = \'"+product_name+"\'") as cur:
            for row in cur:
                product_id = row[0]
        arcpy.AddMessage("Product {0} ({1}) started".format(product_name, product_id))
        # Fetch all reference information from pagination file
        item_list = []
        with arcpy.da.SearchCursor(pagination_file, ["Export_Name"], "PRODUCT_ID = " + str(product_id)) as cur:
            for row in cur:
                item_list.append(row[0])

        # Open the template file
        current_mxd = arcpy.mapping.MapDocument(template_file)

        # Load pagination Layer for re-centering
        df = arcpy.mapping.ListDataFrames(current_mxd)[0]
        lyr = arcpy.MakeFeatureLayer_management(
            pagination_file,
            "temp_pagination").getOutput(0)
        arcpy.mapping.AddLayer(df, lyr, "BOTTOM")
        extent_layer = arcpy.mapping.ListLayers(current_mxd, "temp_pagination", df)[0]
        extent_layer.visible = False
        arcpy.RefreshTOC()
        arcpy.RefreshActiveView()
        # If the page range parameter has any value convert string into list individual of values
        # Has to be strings for find replace against contents of product library to work
        if page_range and page_range.strip():
            raw_list = page_range.split(",")
            clean_list = []
            first_product_parts = []
            last_product_parts = []
            for value in raw_list:
                # If value is a range
                if "-" in value:
                    # A range has been submitted find all the values between start and end
                    split_list = value.split("-")

                    # If the product name contains underscore to separate name from number
                    if "_" in split_list[0]:
                        # The first page split by underscore
                        first_product_parts = split_list[0].split("_")
                        # The last page split by underscore
                        last_product_parts = split_list[-1].split("_")
                        # The first page number (should always be second value in file name, allows for quarters)
                        first_page_number = int(first_product_parts[1])
                        # The last page number plus one to make sure it the whole list is covered
                        # (should always be second value in file name, allows for quarters)
                        last_page_number = int(last_product_parts[1]) + 1

                    else:
                        first_page_number = int(split_list[0])
                        last_page_number = int(split_list[-1]) + 1

                    filled_list = [x for x in range(first_page_number, last_page_number)]
                    # arcpy.AddMessage(filled_list)
                    for a in filled_list:
                        if len(first_product_parts):
                            if a < 1000:
                                # arcpy.AddMessage("Prefix: " + first_product_parts[0])
                                # arcpy.AddMessage("Page: " + str(a).strip())
                                clean_list.append(first_product_parts[0] + "_" + str(a).strip() + "_NW")
                                clean_list.append(first_product_parts[0] + "_" + str(a).strip() + "_NE")
                                clean_list.append(first_product_parts[0] + "_" + str(a).strip() + "_SE")
                                clean_list.append(first_product_parts[0] + "_" + str(a).strip() + "_SW")
                            else:
                                clean_list.append(first_product_parts[0] + "_" + str(a).strip())
                        else:
                            if a < 1000:
                                clean_list.append(str(a).strip() + "_NW")
                                clean_list.append(str(a).strip() + "_NE")
                                clean_list.append(str(a).strip() + "_SE")
                                clean_list.append(str(a).strip() + "_SW")
                            else:
                                clean_list.append(str(a).strip())
                # If value is single entry
                else:
                    if "_" in value:
                        if len(value.split("_")) == 2:
                            page_no = int(value.split("_")[1])
                            arcpy.AddMessage("Page to be run: {0}".format(page_no))
                            if page_no < 1000:
                                # A single product/page number
                                clean_list.append(value.strip() + "_NW")
                                clean_list.append(value.strip() + "_NE")
                                clean_list.append(value.strip() + "_SE")
                                clean_list.append(value.strip() + "_SW")
                            else:
                                clean_list.append(value.strip())
                        if len(value.split("_")) == 3:
                            if value.split("_")[2] in ("NW", "NE", "SE", "SW"):
                                clean_list.append(value.strip())
        else:
            clean_list = item_list

        # Regular expression for product id check/updating
        product_regex = re.compile(r"""PRODUCT_ID\s*=\s*[0-9]*\s""")

        # Regular expression for page number check/updating
        page_no_regex = re.compile(r""" AND PAGE_NO\s*=\s*[0-9]*""")

        # If range was submitted filter the list to include only products in requested series
        if len(clean_list) > 0 and all_pages == "false":
            filtered_item_list = [x for x in clean_list if x in item_list]
        else:
            filtered_item_list = item_list

        arcpy.AddMessage("AOI List Loaded")
        # Loop through the list of products, check each one out, export to PDF and then check it back in
        for i in filtered_item_list:

            # Try each product in turn but don't fail for exception on each one.
            try:
                arcpy.AddMessage("Processing AOI: " + i)

                # Replace definition query page number with mxd number
                filename = i
                pageNumber = filename.split("_")[1]

                layers = arcpy.mapping.ListLayers(current_mxd)
                for lyr in layers:
                    if lyr.supports("DEFINITIONQUERY"):
                        # Annotation Layer is classified as Group Layer
                        if lyr.isGroupLayer or lyr.isFeatureLayer:
                            lyr.definitionQuery = re.sub(product_regex, "PRODUCT_ID = " + str(product_id) + " ",
                                                         lyr.definitionQuery)
                            lyr.definitionQuery = re.sub(page_no_regex, " AND PAGE_NO = " + pageNumber,
                                                         lyr.definitionQuery)
                arcpy.AddMessage("AOI {0} definition queries adjusted (Page No: {1})".format(i, pageNumber))
                # Re-centre using external pagination file
                arcpy.SelectLayerByAttribute_management(extent_layer, "NEW_SELECTION", " \"Export_Name\" = \'" + filename + "\' ")
                df.panToExtent(extent_layer.getSelectedExtent(False))
                arcpy.SelectLayerByAttribute_management(extent_layer, "CLEAR_SELECTION")

                # arcpy.Delete_management(lyr)
                # arcpy.mapping.RemoveLayer(df, extent_layer)
                arcpy.AddMessage("AOI {0} re-centred on extent".format(i))

                # Set full output path for the exported PDF
                outputPath = os.path.join(opath, i + ".pdf")

                # Export to Production PDF
                arcpyproduction.mapping.ExportToProductionPDF(current_mxd,
                                                              outputPath,
                                                              settings_file,
                                                              data_frame="PAGE_LAYOUT",
                                                              resolution=750,
                                                              image_quality="BEST",
                                                              colorspace="CMYK",
                                                              compress_vectors=True,
                                                              image_compression="LZW",
                                                              picture_symbol="VECTORIZE_BITMAP"
                                                              )
                arcpy.AddMessage("AOI {0} from Product {1} Exported".format(i, product_name))

                # Save file to allow any definition query or extent corrections to be persisted
                # current_mxd.saveACopy(os.path.join(opath, os.path.basename(current_mxd.filePath)))
                # current_mxd.save()
                arcpy.env.Workspace = "in_memory"
                arcpy.Delete_management("book_products")
            except Exception as e:
                arcpy.AddError("AOI {0} from Product {1} failed".format(i, product_name))
                arcpy.AddError("Error on line {} {} {}".format(sys.exc_info()[-1].tb_lineno, type(e).__name__, e))
                arcpy.AddError(e)
                continue
        del item_list
        gc.collect()
        # Check in the extension
        arcpy.CheckInExtension("foundation")
        arcpy.env.Workspace = "in_memory"
        arcpy.Delete_management("book_products_view")
        return