'''Model representing a slide and objects on a slide'''

import os
import xml.etree.ElementTree as etree


class slide():

    def __init__(self, num, settings):
        self.settings = settings
        self.slide_num = num
        self.path_on_disk = os.path.join(
            self.settings.temp, "ppt", "slides", "slide" + str(self.slide_num) + ".xml")
        self.notes_path = os.path.join(
            self.settings.temp, "ppt", "notesSlides")
        self.rels_file = os.path.join(
            self.settings.temp, "ppt", "slides", "_rels", "slide" + str(self.slide_num) + ".xml.rels")
        self.objects = []  # ordered list of text & image objects
        self.note_objects = []  # ordered list of text objects from notes
        self.image_assocs = dict()  # dict from rId to image names

    def get_markdown_string(self):
        '''returns the markdown representation of the slide'''
        builder = ""
        for item in self.objects:
            builder = os.linesep + builder + item.get_markdown_string() + os.linesep
        if self.settings.include_notes:
            for item in self.note_objects: ## add each line from notes with prefix if appropriate
                for line in item.get_markdown_string().split(os.linesep):
                    if line == "":
                        continue
                    builder = builder + self.settings.notes_prefix + line + os.linesep
        if self.settings.headings_first: #prefix slide with ## to indicate heading, if appropriate
            builder = os.linesep + "## " + builder.lstrip().lstrip("*")
        else:
            builder = os.linesep + builder.lstrip()
        return builder

    def fetch(self):
        '''Loads a single slide in its entirety, including notes'''
        self.__fetch_rel()
        self.__fetch_notes()
        self.__fetch_content()

    def __fetch_notes(self):
        '''read in slide notes'''
        if os.path.isdir(self.notes_path):  # return, no notes here
            return
        main_schema = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
        root_el = main_schema + "cSld"
        sub_root_slides = main_schema + "spTree"

        textbox_shape = main_schema + "sp"
        table = main_schema + "graphicFrame"
        image = main_schema + "pic"
        for item in etree.parse(self.notes_path).getroot().find(root_el).find(sub_root_slides):
            if item.tag == textbox_shape:
                self.note_objects.append(textObj(item))
            if item.tag == table:
                continue  # ignore for now
        return

    def __fetch_content(self):
        '''read in slide content'''
        main_schema = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
        root_el = main_schema + "cSld"
        sub_root_slides = main_schema + "spTree"

        # only images and text are meaningful in markdown. Things like tables
        # (graphicFrame) are ignored
        textbox_shape = main_schema + "sp"
        image = main_schema + "pic"

        for item in etree.parse(self.path_on_disk).getroot().find(root_el).find(sub_root_slides):
            if item.tag == image:  # add image to slide
                img_id = self.__id_from_pic_node(item)
                img = imageObj(self.settings.file_name, img_id, self.image_assocs[img_id])
                self.objects.append(img)
            if item.tag == textbox_shape:  # add textbox to slide
                self.objects.append(textObj(item))
        return

    def __id_from_pic_node(self, image_node):
        '''Returns the ID of an image from the parent spTree>pic object'''
        sub_node = "{http://schemas.openxmlformats.org/presentationml/2006/main}blipFill"
        blip = '{http://schemas.openxmlformats.org/drawingml/2006/main}blip'
        blip_id_attr = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'

        return image_node.find(sub_node).find(blip).attrib[blip_id_attr]

    def __fetch_rel(self):
        '''Read in slides associations between rId and image/note file'''
        rel_xq = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
        tree_root = etree.parse(self.rels_file).getroot()
        image_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        notes_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"

        for rel in tree_root.findall(rel_xq):
            if rel.attrib["Type"] == image_type:
                self.image_assocs[rel.attrib["Id"]] = os.path.basename(rel.attrib[
                                                                       "Target"])
            if rel.attrib["Type"] == notes_type:
                self.notes_path = os.path.join(
                    self.notes_path, os.path.basename(rel.attrib["Target"]))


class imageObj():
    '''Represents an image object on a slide'''

    def __init__(self, pres_name, ppt_id, file_name):
        '''pres_name is used to prefix image file names to avoid collisions when multiple ppts are in the same dir'''
        self.pptID = ppt_id  # id used to identify resource in pptID
        self.filename = file_name # e.g. image1.png
        self.pres_name = pres_name
        self.alt = ""  # alt text

    def get_markdown_string(self):
        return "![" + self.alt + "](images/" + self.pres_name + "_" + self.filename + ")" + os.linesep


class textObj():
    '''Represents a text/shape object on a slide.'''

    def __init__(self, text_node):
        '''Takes an xml element and builds the text object from it.'''
        self.paras = []  # list of text paras
        self.__build_from_xml_node(text_node)

    def get_markdown_string(self):
        builder = ""
        for para in self.paras:
            if para != "":
                builder = builder + para + os.linesep
        return builder

    def __build_from_xml_node(self, text_node):
        '''returns full textobject from text_node
        text_node: xml object containing text shape/object
        '''
        sub_node = '{http://schemas.openxmlformats.org/presentationml/2006/main}txBody'
        p_node = '{http://schemas.openxmlformats.org/drawingml/2006/main}p'
        r_node = '{http://schemas.openxmlformats.org/drawingml/2006/main}r'
        t_node = '{http://schemas.openxmlformats.org/drawingml/2006/main}t'
        pPR_node = '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr'
        try:
            hasBullet = False
            indent = False
            for para in text_node.find(sub_node).findall(p_node):
                prop_node = para.find(pPR_node)
                span_builder = ""
                if prop_node != None and "lvl" in prop_node.attrib:
                    if not hasBullet:
                        span_builder += os.linesep + "    * "
                        hasBullet = True  # only add extra newline once
                    else:
                        span_builder += "    * "
                    indent = True
                elif indent == False:
                    span_builder += "* "
                for span in para.findall(r_node):
                    span_builder += str(span.find(t_node).text)
                if span_builder != "* " and span_builder.strip() != "":
                    self.paras.append(span_builder)
        except AttributeError:  # empty
            return
        return
