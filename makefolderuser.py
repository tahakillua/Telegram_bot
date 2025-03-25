import os, shutil


class MakeFolderUser:
    def __init__(self, user_name):
        self.user_name = user_name
        self.pathsend = f"./document_send_to_client/{self.user_name}"
        self.pathreceive = f"./file_excel_receive/{self.user_name}"
        self.nameclass = ""
        self.path1 = ""
        self.path2 = ""
        self.path3 = ""
        self.path4 = ""
        self.path5 = ""
        self.path6 = ""
        self.path_picture = None
        self.makefolderusersend()
        self.makefolderuserreceive()
        self.getpathpicture()

    def makefolderusersend(self):
        try:
            os.makedirs(self.pathsend)
        except EOFError:
            pass

    def makefolderuserreceive(self):
        try:
            os.makedirs(self.pathreceive)
        except EOFError:
            pass

    def deletefolderusersned(self):
        try:
            shutil.rmtree(self.pathsend)
        except EOFError:
            pass

    def deletefolderreceive(self):
        try:
            shutil.rmtree(self.pathreceive)
        except EOFError:
            pass

    def getpathpicture(self):
        self.path1 = f"./document_send_to_client/{self.user_name}/{self.nameclass}/picturepercentagedistribution"
        self.path2 = f"./document_send_to_client/{self.user_name}/{self.nameclass}/picturehistogram"
        self.path3 = f"./document_send_to_client/{self.user_name}/{self.nameclass}/pictureboxplot"
        self.path4 = f"./document_send_to_client/{self.user_name}/{self.nameclass}/picturebarchart"
        self.path5 = f"./document_send_to_client/{self.user_name}/{self.nameclass}/picturescatterplot"
        self.path6 = f"./document_send_to_client/{self.user_name}/{self.nameclass}/picturecorrolation"
        self.path_picture = {
            "[picture_percentage_distribution]": self.path1,
            "[picture_histogram]": self.path2,
            "[picture_box_plot]": self.path3,
            "[picture_bar_chart]": self.path4,
            "[picture_scatter_plot]": self.path5,
            "[picture_corrolation]": self.path6
        }

    def getpathclass(self, nameclass):
        try:
            self.nameclass = nameclass
            os.makedirs(self.pathsend + f"/{self.nameclass}")
        except EOFError:
            pass