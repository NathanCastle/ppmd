'''Model representing presentation metadata'''

import os
import xml.etree.ElementTree as etree


class metadata():

    def __init__(self, settings):
        self.title = ""
        self.subject = ""
        self.description = ""
        self.creators = ""
        self.last_modified_by = ""
        self.created = ""
        self.last_modified_at = ""
        self.settings = settings

    def fetch(self):
        '''load metadata from presentation'''
        # open temp_folder_path/docProps/core.xml
        try:
            properties_tree = etree.parse(os.path.join(
                self.settings.temp, "docProps", "core.xml"))
        except FileNotFoundError:  # sometimes properties doesn't exist...
            return
        xmlns_purl = "{http://purl.org/dc/elements/1.1/}"
        xmlns_tr = "{http://purl.org/dc/terms/}"
        xmlns_oo = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
        property_root = properties_tree.getroot()
        self.title = property_root.find(xmlns_purl + "title").text
        self.subject = property_root.find(xmlns_purl + "subject").text
        self.creator = property_root.find(xmlns_purl + "creator").text
        self.description = property_root.find(xmlns_purl + "description").text
        self.last_modified_by = property_root.find(
            xmlns_oo + "lastModifiedBy").text
        self.created = property_root.find(xmlns_tr + "created").text
        self.last_modified_at = property_root.find(xmlns_tr + "modified").text
        return

    def get_markdown_string(self):
        builder = ""
        builder += "title: " + str(self.title) + os.linesep
        builder += "subject: " + str(self.subject) + os.linesep
        builder += "description: " + str(self.description) + os.linesep
        builder += "creators: " + str(self.creators) + os.linesep
        builder += "last_modified_by: " + \
            str(self.last_modified_by) + os.linesep
        builder += "created: " + str(self.created) + os.linesep
        builder += "last_modified_at: " + \
            str(self.last_modified_at) + os.linesep
        return builder
