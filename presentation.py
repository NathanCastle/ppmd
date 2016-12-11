import zipfile
import os
import xml.etree.ElementTree as etree
import shutil
import slide
import metadata


class pres_settings():

    def __init__(self, slide_sep, headings_first, filepath, includenotes, notesprefix):
        self.slide_sep = slide_sep
        self.headings_first = headings_first
        self.file_path = filepath
        self.base_folder = os.path.dirname(filepath)
        self.file_name = os.path.splitext(os.path.basename(filepath))[0]
        self.temp = os.path.join(self.base_folder, self.file_name)
        self.include_notes = includenotes
        self.notes_prefix = notesprefix


class presentation():

    def __init__(self, settings):
        self.settings = settings
        self.meta = metadata.metadata(self.settings)
        self.slides = []

    def fetch(self):
        '''Begin reading in content'''
        self.__unzip()
        self.meta.fetch()
        self.__enumerate_slides()
    
    def write_output(self):
        '''Write content to disk, then clean up'''
        self.__prep_dir()
        self.__prep_images()
        with open(os.path.join(self.settings.base_folder, self.settings.file_name + ".md"), "w") as output:
            output.write(self.get_markdown_string())
        self.__cleanup()

    def __unzip(self):
        '''unzips the presentation, writing the unzipped folder to disk'''
        zipped = zipfile.ZipFile(os.path.join(
            self.settings.base_folder, self.settings.file_name + ".pptx"), 'r')
        # extract pptx to directory (ooxml files are secretly zipped dirs)
        zipped.extractall(self.settings.temp)
        zipped.close()

    def __cleanup(self):
        '''Need to delete the intermediate unzipped folder'''
        shutil.rmtree(self.settings.temp)

    

    def __prep_dir(self):
        '''prep any items needed for output'''
        # e.g. need to create images folder
        image_dir = os.path.join(self.settings.base_folder, "images")
        if not os.path.exists(image_dir):
            os.makedirs(image_dir)

    def __prep_images(self):
        '''Copy images to output directory'''
        ppt_media_path = os.path.join(self.settings.temp, "ppt", "media")
        output_img_path = os.path.join(self.settings.base_folder, "images")
        images = os.listdir(ppt_media_path)
        for img in images:
            shutil.copy2(os.path.join(ppt_media_path, img), os.path.join(
                output_img_path, self.settings.file_name + "_" + img))
        return

    def __enumerate_slides(self):
        for item in os.listdir(os.path.join(self.settings.temp, "ppt", "slides")):
            if item.startswith("slide"):
                xml_slide = slide.slide(
                    int(os.path.splitext(item)[0][5:]), self.settings)
                self.slides.append(xml_slide)
                xml_slide.fetch()
        return

    def get_markdown_string(self):
        builder = ""
        builder += self.meta.get_markdown_string()
        builder += os.linesep
        self.slides.sort(key=lambda x: x.slide_num)
        for slide_item in self.slides:
            builder += self.settings.slide_sep + os.linesep + \
                slide_item.get_markdown_string() + os.linesep
        return builder
