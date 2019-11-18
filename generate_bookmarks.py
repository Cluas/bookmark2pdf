# coding:utf-8
import argparse
import os

from PyPDF2 import PdfFileReader as reader
from PyPDF2 import PdfFileWriter as writer


class PDFHandleMode(object):
    """
    处理PDF文件的模式
    """

    COPY = "copy"
    NEWLY = "newly"


class MyPDFHandler(object):
    """
    封装的PDF文件处理类
    """

    def __init__(self, pdf_file_path, mode=PDFHandleMode.COPY):
        """
        用一个PDF文件初始化
        :param pdf_file_path: PDF文件路径
        :param mode: 处理PDF文件的模式，默认为PDFHandleMode.COPY模式
        """
        self.__pdf = reader(pdf_file_path)
        self.file_name = os.path.basename(pdf_file_path)
        self.metadata = self.__pdf.getXmpMetadata()
        self.doc_info = self.__pdf.getDocumentInfo()
        self.pages_num = self.__pdf.getNumPages()

        self.__writeable_pdf = writer()
        if mode == PDFHandleMode.COPY:
            self.__writeable_pdf.cloneDocumentFromReader(self.__pdf)
        elif mode == PDFHandleMode.NEWLY:
            for idx in range(self.pages_num):
                page = self.__pdf.getPage(idx)
                self.__writeable_pdf.insertPage(page, idx)

    def save2file(self, new_file_name):
        """
        将修改后的PDF保存成文件
        :param new_file_name: 新文件名，不要和原文件名相同
        :return: None
        """
        # 保存修改后的PDF文件内容到文件中
        with open(new_file_name, "wb") as out:
            self.__writeable_pdf.write(out)

    def add_one_bookmark(self, title, page, parent=None, color=None, fit="/Fit"):
        """
        往PDF文件中添加单条书签，并且保存为一个新的PDF文件
        :param str title: 书签标题
        :param int page: 书签跳转到的页码，表示的是PDF中的绝对页码，值为1表示第一页
        :param parent: A reference to a parent bookmark to create nested bookmarks.
        :param tuple color: Color of the bookmark as a red, green, blue tuple from 0.0 to 1.0
        :param str fit: 跳转到书签页后的缩放方式
        :return: Bookmark
        """
        # 为了防止乱码，这里对title进行utf-8编码
        return self.__writeable_pdf.addBookmark(title, page - 1, parent=parent, color=color, fit=fit)

    def add_bookmarks(self, bookmarks):
        """
        批量添加书签
        :param bookmarks: 书签元组列表，其中的页码表示的是PDF中的绝对页码，值为1表示第一页
        :return: None
        """
        parent = None
        for title, page in bookmarks:
            if title.startswith("    -"):
                self.add_one_bookmark(title.split("-")[1].strip(), page, parent)
            else:
                parent = self.add_one_bookmark(title, page)

    @staticmethod
    def read_bookmarks_from_txt(txt_file_path, page_offset=0):
        """
        从文本文件中读取书签列表
        文本文件有若干行，每行一个书签，内容格式为：
        书签标题 页码
        注：中间用空格隔开，页码为1表示第1页
        :param txt_file_path: 书签信息文本文件路径
        :param page_offset: 页码便宜量，为0或正数，即由于封面、目录等页面的存在，在PDF中实际的绝对页码比在目录中写的页码多出的差值
        :return: 书签列表
        """
        bookmarks = []
        with open(txt_file_path, "r") as fin:
            for line in fin:
                line = line.rstrip()
                if not line:
                    continue
                # 以'@'作为标题、页码分隔符
                try:
                    title = line.split("@")[0].rstrip()
                    page = line.split("@")[1].strip()
                except IndexError as msg:
                    continue
                # title和page都不为空才添加书签，否则不添加
                if title and page:
                    try:
                        page = int(page) + page_offset
                        bookmarks.append((title, page))
                    except ValueError as msg:
                        print (msg)

        return bookmarks

    def add_bookmarks_by_read_txt(self, txt_file_path, page_offset=0):
        """
        通过读取书签列表信息文本文件，将书签批量添加到PDF文件中
        :param txt_file_path: 书签列表信息文本文件
        :param page_offset: 页码便宜量，为0或正数，即由于封面、目录等页面的存在，在PDF中实际的绝对页码比在目录中写的页码多出的差值
        :return: None
        """
        bookmarks = self.read_bookmarks_from_txt(txt_file_path, page_offset)
        self.add_bookmarks(bookmarks)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='pdf bookmark generation tool.')
    parser.add_argument('--input', dest='i', action='store',
                        help='origin pdf filename.')
    parser.add_argument('--txt', dest='t', action='store',
                        help='txt of bookmarks.')
    parser.add_argument('--output', dest='o', action='store',
                        help='save to filename.')
    args = parser.parse_args()
    pdf_handler = MyPDFHandler(args.i, mode=PDFHandleMode.NEWLY)
    print("reading origin pdf file success...")
    pdf_handler.add_bookmarks_by_read_txt(args.t, page_offset=0)
    print("parse bookmark success...")
    pdf_handler.save2file(args.o)
    print("save pdf with bookmark success...")
