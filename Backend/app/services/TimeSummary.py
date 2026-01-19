import pandas as pd
import io


def get_summary(file_content: bytes):
    # 1. 讀取原始檔案
    source_file = pd.read_excel(file_content, sheet_name=None)
    df_list = []
    output = io.BytesIO()

    # 建立 Excel 寫入器
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, df in source_file.items():
            if '月' in name:
                # 清理全空欄位
                clean_df = df.dropna(axis=1, how='all').copy()

                if clean_df.empty or clean_df.shape[1] <= 1:
                    clean_df.to_excel(writer, sheet_name=name, index=False)
                    continue
                
                try:
                    # 先找到日期 '1' 的原始位置 (這是計算的關鍵基準)
                    # 我們先計算數值，暫不插入欄位，避免 index 跑掉
                    col_names = list(clean_df.columns.map(str))
                    col_index_1 = col_names.index('1')
                    
                    # 計算數值
                    s_counts = clean_df.iloc[:, col_index_1:].count(axis=1)
                    s_hours = clean_df.iloc[:, col_index_1:].sum(axis=1)
                    t_fees = s_counts * 50
                    
                    # 處理欄位插入 (Logic: 為了避免重複插入報錯，先 drop 再 insert)
                    cols_to_add = [
                        (2, '交通費', t_fees),
                        (2, '服務時數', s_hours),
                        (4, '服務次數', s_counts)
                    ]
                    
                    for pos, col_name, data in cols_to_add:
                        if col_name in clean_df.columns:
                            clean_df.drop(columns=[col_name], inplace=True)
                        clean_df.insert(pos, col_name, data)

                    # 寫入分頁
                    clean_df.to_excel(writer, sheet_name=name, index=False)
                    
                    # 為了總表，我們只保留需要的欄位，減少記憶體消耗
                    summary_part = clean_df[['志工姓名', '服務時數', '服務次數', '交通費']].copy()
                    df_list.append(summary_part)

                except ValueError:
                    print(f"分頁 {name} 找不到日期 '1'，已跳過。")
            
            elif name != '總表':
                # 非月份分頁直接複製
                df.to_excel(writer, sheet_name=name, index=False)

        # 2. 製作總表 (Summary Logic)
        if df_list:
            df_all = pd.concat(df_list)
            # 使用 groupby 加總
            df_summary_all = df_all.groupby('志工姓名', as_index=False).sum()
            
            # 重新命名欄位以便區分
            df_summary_all.rename(columns={'服務時數': '總累積時數'}, inplace=True)
            
            # 排名邏輯
            df_summary_all['排名'] = df_summary_all['總累積時數'].rank(ascending=False, method='min').astype(int)
            
            # 排序：按排名排
            df_summary_all = df_summary_all.sort_values('排名')
            df_summary_all.to_excel(writer, sheet_name='總表', index=False)
            
    # Prepare the buffer for reading
    output.seek(0)
    
    return output
            

if __name__ == '__main__':
    # 設定路徑
    input_file = r"D:\志工時數結算(114總整理).xlsx"
    output_file = r"D:\志工時數結算_已處理.xlsx"
    
    get_summary(input_file)
    print('處理完成！檔案已存於 D 槽。')