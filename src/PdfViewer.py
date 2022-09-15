from typing import Tuple

import fitz


class PdfViewer(object):
    __slot__ = ('name', 'doc', 'page_count', 'dlist_tab')

    def __init__(self) -> None:
        self.name = ''
        self.page_count = 0
    
    def open(self, name:str) -> None:
        if self.name != name:
            self.name = name
            self.doc = fitz.open(self.name)
            self.page_count = self.doc.page_count
            self.dlist_tab = [None] * self.page_count

    def get_page(self, pno:int) -> Tuple[int, bytes]:
        if pno > self.page_count-1:
            pno = 0
        elif pno < 0:
            pno = self.page_count-1
        dlist = self.dlist_tab[pno]
        if dlist is None:
            self.dlist_tab[pno] = self.doc[pno].get_displaylist()
            dlist = self.dlist_tab[pno]
            
        pix = dlist.get_pixmap(alpha=False)
        return pno, pix.tobytes()
