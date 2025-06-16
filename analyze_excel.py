import os
import pandas as pd
import numpy as np
import colorsys
import matplotlib.pyplot as plt
import json
import random
import ipdb

random.seed(1234567)

class AutoMarket():
    def __init__(self, working_dir, dump_vis=False):
        self.working_dir = working_dir
        self.res_dir = f'{self.working_dir}/res'
        os.makedirs(self.res_dir, exist_ok=True)
        self.dump_vis = dump_vis
        self.order_cust = ['通威', '晶科', '晶澳', '天合', '阿特斯', '捷泰', '正泰', '隆基', '英发', '中润', '爱旭']
        self.header_prefix = ['客户', '基地', '负责人', '实际开线数', '日用量（KG/线)', '总数核对', '备注', '电池尺寸',
                                '单片湿重（mg）', '日用量（KG/线）']
        self.products = ['正面副栅', '背面副栅', '正面主栅', '背面主栅']
        self.customers = []
        self.competitors = None

    def split_excel(self, df):
        '''
        把src_dataframe拆分为两个
        '''
        r,c = df.shape
        r_split = []
        i = 1
        while i<r:
            while i<r and (not pd.isna(df.iloc[i,1])):
                i += 1
            if i<r and pd.isna(df.iloc[i+1,1]):
                r_split.append(i-1)
            while i<r and pd.isna(df.iloc[i,1]):
                i+=1
            if i<r and (not pd.isna(df.iloc[i+1, 1])):
                r_split.append(i)
        # print(r_split)
        return df.iloc[0:r_split[0]+1], df.iloc[r_split[1]:r+1]

    def unmerge(self, df):
        '''
        填充因合并单元格导致的缺失label
        确保key唯一:
            将第一行和第二行的item合并为新item
            将客户 基地 负责人合并为新item
        '''
        # 填充表头(第一行)
        headers = df.columns.tolist()
        assert not pd.isna(headers[0]), '第一列的表头不能为空'
        i = 0
        while i<len(headers):
            while i<len(headers) and ('Unnamed:' not in headers[i]):
                i += 1
            while i<len(headers) and ('Unnamed:' in headers[i]):
                headers[i] = headers[i-1]
                i += 1
        df.columns = headers

        # 从第一行中提取competitors
        competitors = headers.copy()
        i = 0
        while i<len(competitors):
            if competitors[i] in self.header_prefix:
                competitors.pop(i)
            else:
                i += 1

        # 合并第一行和第二行为新索引
        self.company_product = headers
        product = df.iloc[0].tolist()
        assert len(product)==len(headers), '第一行与第二行列数不相等'
        for i in range(len(product)):
            if pd.isna(product[i]):
                continue
            self.company_product[i] = f'{headers[i]}_{product[i]}'
        # reset headers
        df.columns = self.company_product
        # 删除第二行
        df = df.iloc[1:]
        
        # self.competitors去重
        competitors = list(set(competitors))
        if not self.competitors:
            self.competitors = competitors
        print(self.competitors)
        assert self.competitors==competitors, 'competitor conflict'

        # 填充第一列
        customers = df['客户'].tolist()
        
        print(len(customers))
        i = 0
        while i<len(customers):
            while i<len(customers) and (not pd.isna(customers[i])):
                i += 1
            while i<len(customers) and i>0 and pd.isna(customers[i]):
                customers[i] = customers[i-1]
                i += 1
            i += 1
        df['客户'] = customers
        self.customers = list(set(customers+self.customers))

        # 重置第一列
        jidi = df['基地'].tolist()
        fuzeren = df['负责人'].tolist()
        assert len(jidi)==len(fuzeren), '基地和负责人数量对不齐'
        assert len(jidi)==len(customers), '基地和客户数量对不齐'
        idx_new = customers
        for i in range(len(idx_new)):
            idx_new[i] = f'{customers[i]}_{jidi[i]}_{fuzeren[i]}'
        df.columns = ['客户_基地_负责人']+df.columns.tolist()[1:]
        df['客户_基地_负责人'] = idx_new
        # 删除基地 负责人列
        del df['基地']
        del df['负责人']

        return df

    def analyze(self, df, json_name='main'):
        # info 保存了各客户-各产品-各友商的线数
        self.cust_dfs = dict()
        info = dict()
        for cust in self.customers:
            info[cust] = dict()
            filtered_df = df[df.iloc[:, 0].astype(str).str.contains(cust, na=False)]
            info[cust]['实际开线数'] = int(pd.to_numeric(filtered_df['实际开线数'], errors='coerce').sum())
            # ipdb.set_trace()
            for prod in self.products:
                info[cust][prod] = dict()
                for compe in self.competitors:
                    company_product = f'{compe}_{prod}'
                    sum_company_product = filtered_df[company_product].sum()
                    info[cust][prod][compe] = sum_company_product
                if self.dump_vis:
                    self.draw_pie(info[cust][prod].keys(), info[cust][prod].values(), title=f'{cust}_{prod}')
            
        self.dump_analyze_results(info, json_name) 

        # 行业各production友商share分析
        self.prod_share = dict()
        for prod in self.products:
            self.prod_share[prod] = dict()
            for compe in self.competitors:
                company_product = f'{compe}_{prod}'
                columns_with_comp_prod = [col for col in df.columns if company_product in col.lower()]
                filtered_df = df[columns_with_comp_prod]
                sum_company_product = filtered_df[company_product].sum()
                self.prod_share[prod][compe] = sum_company_product
            if self.dump_vis:
                self.draw_pie(info[cust][prod].keys(), info[cust][prod].values(), title=f'{prod}_大客户')
        df_prod_share = pd.DataFrame(self.prod_share)
        # ipdb.set_trace()
        return info
            


    def dump_analyze_results(self, info, json_name='main.json'):
        # 将统计结果保存为json,方便后续回溯读取
        with open(f'{self.res_dir}/{json_name}.json', 'w', encoding='utf-8') as f:
            json.dump(info, f, ensure_ascii=False, indent=4) 



    def prepare(self, df):
        '''
        切分df为main与other
        填充第一行第一列的nan
        '''
        # df_main, df_other = self.split_excel(df)
        df_main = self.unmerge(df)
        return df_main, None


    def forward(self, sheet='客户I'):
        print('\033[1;32m---- prepare data ----\033[0m')
        # data prepare
        dfs = []
        for f in os.listdir(self.working_dir):
            if not f.endswith('.xlsx'):
                continue
            excel_path = os.path.join(self.working_dir, f)
            print(f'\033[1;32m---- process {excel_path} ----\033[0m')
            # read main customer
            df_main = pd.read_excel(excel_path, sheet_name=sheet)
            self.excel_path = excel_path
            df_main, _ = self.prepare(df_main)
            dfs.append(df_main)

        df_main = pd.concat(dfs, ignore_index=True)
        # df_main.to_excel(f'{self.res_dir}/{sheet}.xlsx', index=False)
        print(self.competitors)
        print('\033[1;32m---- analyze data ----\033[0m')
        info_main = self.analyze(df_main, json_name=sheet)
        print('\033[1;32m---- finished ----\033[0m')

        # adjust order of cust in info
        info_main = self.adjust_order_json(info_main)

        return info_main, df_main

    def adjust_order_json(self, info):
        new_info = dict()
        for c in self.order_cust:
            if c in info:
                new_info[c] = info.pop(c)
        new_info.update(info)
        return new_info







if __name__ == "__main__":
    working_dir = f'./data/20250602'
    market = AutoMarket(working_dir)
    _, _ = market.forward()
            


