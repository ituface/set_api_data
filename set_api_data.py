# coding:utf-8
import xlrd, xlwt
import re,time

class setApiData():
    data_list = []
    read_excel_data = xlrd.open_workbook('../gedai_jiekou/excel/CaseData.xlsx')
    write_excel_data = xlwt.Workbook()
    input_excel_sheet = write_excel_data.add_sheet('output')

    '''
    将exlcel中数据转成list
    '''

    def get_excel_data_to_list(self):
        table = self.read_excel_data.sheets()[1]
        nrows = table.nrows
        data_parameter = []
        data_positive = []
        data_negative = []
        get_api_data = table.row_values(0)[0]
        key = '(.*?)/'
        api_name = re.search(key, get_api_data).group(1)
        api_address = get_api_data.lstrip(api_name)
        for nrows_i in range(nrows):
            if nrows_i <= 1: continue
            data_parameter.append(table.row_values(nrows_i)[1])
            data_positive.append(table.row_values(nrows_i)[2])
            data_negative.append(table.row_values(nrows_i)[3].split('&'))

        return [self.create_api_data(self,data_parameter, data_positive, data_negative), api_name, api_address]

    '''
    创建测试用例，将所有测试用例以list形式返回
    '''

    def create_api_data(self, data, data1, data2):

        '''
        形成正向测试用例 和 字段缺失用例
        '''

        alone_data = '{'
        length = len(data)
        deficiency_list = []  # 缺失字段用例
        for out_i in range(length):

            alone_data = alone_data + '"%s"' % data[out_i] + ':' + "%s"%data1[out_i] + ','
            if out_i == length - 1:
                alone_data = alone_data.rstrip(',')
                alone_data += '}'

            alone_deficiency_data = '{'
            for inner_i in range(length):
                if out_i == inner_i:
                    continue
                alone_deficiency_data = alone_deficiency_data + '"%s"' % data[inner_i] + ':' + "%s"%data1[
                    inner_i] + ','
            alone_deficiency_data = alone_deficiency_data.rstrip(',') + '}'
            deficiency_list.append([alone_deficiency_data, '反向测试用例:缺失字段--->%s' % data[out_i]])

        self.just_data=[alone_data, '正向测试用例']#将正向用例保存下来，添加到用例最后
        self.data_list.extend(deficiency_list)

        '''
        形成反向测试用例
        '''
        if [x for x in data2 if x != ['']] == []:  # 如果没有反向测试用例直接将正向返回
            return self.data_list
        for out_k, out_v in enumerate(data2):  # 遍历反向测试最外层
            for center in out_v:  # 遍历反向测试用例的每一条
                alone_data = '{'
                alone_str = ''
                for k, v in enumerate(data):  # 遍历字段名
                    if k == out_k:
                        alone_data = alone_data + '"%s"' % v + ':' + '%s' % center + ','
                        alone_str = '反向测试用例:%s----->%s' % (v, center)
                        continue
                    alone_data = alone_data + '"%s"' % v + ':' + '%s' % data1[k] + ','
                alone_data = alone_data.strip(',') + '}'
                self.data_list.append([alone_data, alone_str])
        self.data_list.append(self.just_data)
        return self.data_list

    '''
    将生成的数据保存到excel中
    '''
    @classmethod
    def save_data_excel(cls):
        inuput_excel_data, api_name, api_address = cls.get_excel_data_to_list(cls)
        for index, value in enumerate(inuput_excel_data):
            cls.input_excel_sheet.write(index, 0, index + 1)
            cls.input_excel_sheet.write(index, 1, api_name)
            cls.input_excel_sheet.write(index, 2, api_address)
            cls.input_excel_sheet.write(index,5,1)
            for inner in range(len(value)):
                cls.input_excel_sheet.write(index, 3 + inner, value[inner])
        cls.write_excel_data.save('../gedai_jiekou/excel/interfaceDatass.xls')

if __name__ == "__main__":
  setApiData.save_data_excel()

