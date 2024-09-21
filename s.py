import pandas as pd
from xdbSearcher import XdbSearcher

# 读取Excel文件
def read_ip_from_excel(file_path):
    df = pd.read_excel(file_path)  # 确保Excel文件的路径正确
    ip_list = df['IP'].tolist()  # 假设Excel文件中的IP列名称是“IP”
    return ip_list

# 执行IP查询并返回结果
def search_ip(ip_list):
    # 1. 创建查询对象
    dbPath = "D:\\ip2region\\data\\ip2region.xdb"  # 确保ip2region数据库的路径正确
    searcher = XdbSearcher(dbfile=dbPath)
    
    result_list = []
    
    # 2. 批量查询
    for ip in ip_list:
        try:
            region_str = searcher.searchByIPStr(ip)
            result_list.append((ip, region_str))
        except Exception as e:
            result_list.append((ip, f"Error: {str(e)}"))
    
    # 3. 关闭查询器
    searcher.close()
    
    return result_list

# 保存结果到Excel
def save_to_excel(result_list, output_file):
    result_df = pd.DataFrame(result_list, columns=['IP', 'Location'])
    result_df.to_excel(output_file, index=False)

# 主函数，执行流程
def main():
    input_file = 'test.xlsx'  # 输入的Excel文件路径
    output_file = 'ip_query_results.xlsx'  # 输出结果保存路径
    
    # 读取Excel中的IP地址
    ip_list = read_ip_from_excel(input_file)
    
    # 查询IP地址的归属地信息
    results = search_ip(ip_list)
    
    # 保存查询结果到新的Excel文件
    save_to_excel(results, output_file)
    print(f"查询结果已保存到 {output_file}")

if __name__ == "__main__":
    main()
