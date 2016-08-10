# ecoding=utf-8
# Author: 翁彦彬 | Sven_Weng
# Email : diandianhanbin@gmail.com
# Date : 2016-08-10
import xlrd
import time
import xlwt
import sys

reload(sys)
sys.setdefaultencoding("utf-8")


def openExcel(url, index):
	"""
	打开Excel文件(仅支持xls格式的Excel文档),返回一个sheet控制
	:param url: str, 文件的地址
	:param index: int, sheet的序号
	:return: sheet对象
	"""
	data = xlrd.open_workbook(url, formatting_info=True)
	table = data.sheet_by_index(index)
	return table


def getSingleCell(table, x, y):
	"""
	获取单元格的内容
	:param table: sheet对象
	:param x: int, 横坐标
	:param y: int, 纵坐标
	:return: Str, 单元格内容
	"""
	return table.cell(x, y).value


def getRowCell(table, index):
	"""
	获取整列的数据,单独去除了单行后的空元素,单行前的空元素会自动获取所在列之前行的数据
	:param table:sheet对象
	:param index:int, 行号
	:return:list, 整列的数据
	"""
	row = []
	status = 0
	for x in range(table.ncols):
		if getSingleCell(table, index, x) == '' and status == 0:
			row.append(findCell(table, index, x))
		elif getSingleCell(table, index, x) == '' and status == 1:
			break
		else:
			row.append(getSingleCell(table, index, x))
			status = 1
	return row


def findCell(table, rowindex, colindex):
	"""
	已所在列之前行的内容来填充单元格内容
	:param table: sheet对象
	:param rowindex: int, 行号
	:param colindex: int, 列号
	:return: str, 填充之后单元格的内容
	"""
	while 1:
		cell = getSingleCell(table, rowindex, colindex)
		if cell == '':
			rowindex -= 1
		else:
			return cell


def cancleLevel(arr):
	"""
	删除由于Xmind导出后有可能会取到首行Level内容的内容
	:param arr: list, 需要处理的数组
	:return: list, 处理过后的数组
	"""
	cancle = []
	for i, x in enumerate(arr):
		if 'Level' in x:
			cancle.append(x)
	for x in cancle:
		arr.remove(x)
	return arr


def readMain(table):
	"""
	读取当前文件的所有数据,每行的数据以列表形式装在列表中
	:param table: sheet对象
	:return: list, 每行数据的总列表
	"""
	arr = [cancleLevel(getRowCell(table, x)) for x in range(table.nrows)]
	return arr


def writeExcel(table):
	"""
	写入Excel
	:param table: sheet对象
	:return: None
	"""
	title = [u'路径', u'用例编号', u'测试名称', u'描述', u'前置条件', u'步骤描述', u'预期结果', u'优先级', u'用例类型', u'用例属性', u'设计人', u'设计日期']
	f = xlwt.Workbook()
	sheet1 = f.add_sheet(u'TestCase', cell_overwrite_ok=True)
	# 书写每行的内容
	for i, rows in enumerate(readMain(table)):
		sheet1.write(i, 0, '/'.join(rows))
		sheet1.write(i, 10, u'翁彦彬/wengyb')
		sheet1.write(i, 11, time.strftime('%Y-%m-%d'))

	# 书写标题行
	for x in range(len(title)):
		sheet1.write(0, x, title[x])

	f.save('TestCase.xls')


if __name__ == '__main__':
	table = openExcel(u'TestReq.xls', 0)
	# getSingleCell(table, 0, 0)
	# print getRowCell(table, 6)
	# print cancleLevel(getRowCell(table, 6))
	# findCell(table, 5, 1)
	# main(table)
	writeExcel(table)
