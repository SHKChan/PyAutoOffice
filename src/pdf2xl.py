import threading
import win32file

import openpyxl
import pdfplumber


from mylogger import LOGGER


class Pdf2Xl(threading.Thread):
    __slots__ = ('pdf_list', 'xl_path', 'sheet_name',
                 'data', 'exit_code', 'progress')

    title_xl = {'Date': 0, 'Customer': 1, 'PO#1': 2, 'Type': 3, 'PO#2': 4,
                'Description': 5, 'Qty': 6, 'Unit Price': 7, 'Total': 8, 'Customer req. ETD': 9}
    mapping_pdf2xl = {'Item': 3, 'Description': 5,
                      'Qty': 6, None: None, 'Rate': 7, 'Amount': 8}
    table_width = 10

    title_pdf = ['Item', 'Description', 'Qty', None, 'Rate', 'Amount']

    def __init__(self, pdf_list: str, xl_path: str, sheet_name: str = '') -> None:
        self.pdf_list = pdf_list
        self.xl_path = xl_path
        self.sheet_name = sheet_name
        self.data = list()
        self.exit_code = None
        self.progress = 0   # Max value: 100%
        # 创建多线程 设置以保护模式启动，即主线程运行结束，子线程也停止运行
        super().__init__()
        self.setDaemon(True)

    def run(self):
        # Variable that stores the exception, if raised by someFunction
        self.exit_code = None
        try:
            self.convert()
        except Exception as e:
            self.exit_code = e
            LOGGER.wt()

    def join(self):
        threading.Thread.join(self)
        # Since join() returns in caller thread
        # we re-raise the caught exception
        # if any was caught
        if self.exit_code:
            raise self.exit_code

    def convert(self) -> None:
        if not self.is_file_used(self.xl_path, win32file.GENERIC_WRITE):
            self.rd_from_pdf()  # account for 90% in progress
            self.wt_to_xl()  # account for 10% in progress
            self.exit_code = 0
        else:
            self.exit_code = 1
            raise Exception('Excel file is opened!')

    def rd_from_pdf(self) -> None:
        total_last = 0
        len_pdfs = len(self.pdf_list)
        for pdf in self.pdf_list:
            with pdfplumber.open(pdf) as f:
                # 提取PDF中的一页
                for page in f.pages:
                    Date = ''
                    PONo = ''
                    im = page.to_image(resolution=150)
                    # 在页面内,逐表格筛选数据
                    for table in page.extract_tables():
                        # 先找到标题栏标识Date和'P.O. No.'
                        if (Date == '' and ['Date', 'P.O. No.'] in table):
                            # 若已经找到标题栏标识Date,记录Date和P.O.No.
                            Date = table[1][0]
                            PONo = table[1][1]
                        # 已经找到标题栏标识Item, 格式化数据为xl排版
                        if (self.title_pdf in table):
                            self.format_data(table)

                    # 补填Date和P.O No.
                    len_data = len(self.data)
                    for i in range(total_last, len_data):
                        self.data[i][self.title_xl['Date']] = Date
                        self.data[i][self.title_xl['PO#1']] = PONo
                        self.data[i][self.title_xl['PO#2']] = PONo
                    total_last = len_data

            self.progress += 1/len_pdfs*50

    def wt_to_xl(self) -> None:
        # 打开指定路径的Excel文件
        wb = openpyxl.load_workbook(self.xl_path)
        ws_names = wb.get_sheet_names()
        if self.sheet_name == '':
            ws = wb.get_sheet_by_name(ws_names[0])
        else:
            ws = wb.get_sheet_by_name(self.sheet_name)

        # 找第一个空行
        for row in range(1, ws.max_row+2):
            for col in range(1, self.table_width+1):
                if ws.cell(row, col).value == None:
                    first_row2wt = row
                    break

        # 将数据写进sheet表单中的第一个空行
        len_data = len(self.data)
        for row in range(len_data):
            for col in range(10):
                ws.cell(row+first_row2wt, col+1).value = self.data[row][col]
            self.progress += 1/len_data*50

         # 保存文件
        wb.save(self.xl_path)
        wb.close()

    def format_data(self, table: list) -> None:
        # 已经找到标题栏标识Item
        if (self.title_pdf in table):
            len_table = len(table)
            i = 1
            while (i < len_table):
                col = table[i]
                # 避免尾部的空行影响
                if col[2] != '' and col[2] != None:
                    len_col = len(col)
                    # 已经找到标题栏标识Item,在二维矩阵中新增一行
                    self.data.append([])
                    for n in range(self.table_width):
                        self.data[-1].append('')

                    # 记录数据项
                    for j in range(len_col):
                        # 忽略None
                        if col[j] in ['', None]:
                            continue
                        if col[j] == 'Parts':
                            col[j] = 'Part'
                        index = self.mapping_pdf2xl[self.title_pdf[j]]
                        self.data[-1][index] = col[j]
                    # 找出当前item所占的行数(即知道下一item出现为止)
                    index = self.mapping_pdf2xl['Description']
                    description = [self.data[-1][index]]
                    for k in range(i+1, len_table):
                        col_temp = table[k]
                        if (col_temp[0] == ''):
                            # 将属于本item的其余Description继续录入
                            description.append(col_temp[1])
                        else:
                            # 直接跳跃至下一item的对应行
                            self.data[-1][index] = '\n'.join(description)
                            i = k
                            break
                else:
                    i += 1

    def is_file_used(self, file: str, type: int) -> bool:
        # xl文件是否可写
        try:
            vHandle = win32file.CreateFile(self.xl_path,    # 文件名

                                           type,  # 访问对象的类型。应用程序可以获得读访问、写访问、读写访问或设备查询访问

                                           0,   # 指定如何共享对象的位标志集,如果dwShareMode为0，则不能共享该对象。

                                           None,    # 安全属性，或者没有，为None

                                           win32file.OPEN_EXISTING,  # 指定对存在的文件执行哪个操作，以及在文件不存在时执行哪个操作

                                           win32file.FILE_ATTRIBUTE_NORMAL,  # 文件的属性

                                           None)  # 指定对模板文件具有GENERIC_READ访问权限的句柄
            return int(vHandle) == win32file.INVALID_HANDLE_VALUE
        except:
            return True
        finally:
            try:
                win32file.CloseHandle(vHandle)
            except:
                pass
