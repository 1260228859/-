# -*- coding: utf-8 -*-
# @Time         : 18-9-18 下午12:17
# @Author       : EvanKao
# @File         : export_excel.py
# @Description  : 表格导出类

from django.shortcuts import HttpResponse
import xlwt
from io import BytesIO
from urllib import parse


class ExportExcel:

    def __init__(self, excel_name, header, data, labels=None):
        """
        初始化
        :param excel_name: 表名 :str
        :param header: 头部名称 :list
        :param data: 表格数据 :dict
        :param lables: 表格数据对应中文注释 :dict
        """
        self.excel_name = excel_name
        self.wb = xlwt.Workbook(encoding='utf-8')
        self.header = header
        self.labels = labels if labels else {}
        self.all_data = data

    def wirte_data(self):
        """
        导出
        :return:
        """
        data = [[self.labels.get(column, column) for column in self.header]]
        for item in self.all_data:
            # for item_val in item.values():
            data.append([item.get(column) for column in self.header])
        sheet_prd = self.wb.add_sheet('sheet1')
        for row_index, row in enumerate(data):
            for column_index, column in enumerate(row):
                column = column if column else 0
                sheet_prd.write(row_index, column_index, column)

    def get_excel_name(self):
        return parse.quote(self.excel_name)

    def __call__(self):
        self.response = HttpResponse(content_type='application/vnd.ms-excel')
        execl_name = self.get_excel_name()
        self.response['Content-Disposition'] = 'attachment;filename={0}.xls'.format(execl_name)
        self.wb = xlwt.Workbook(encoding='utf-8')
        self.wirte_data()

        output = BytesIO()
        self.wb.save(output)
        output.seek(0)
        self.response.write(output.getvalue())
        return self.response
