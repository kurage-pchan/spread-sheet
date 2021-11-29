worksheet_len = [ pd.DataFrame((workbook.get_worksheet(j)).get_all_values()) for j in range( len( worksheet_list ) ) ]
list_len = [worksheet_len[p].iloc[1].to_list() for p in range( len( worksheet_list ))]
ave = 0
for j in list_len :#すべてのワークシートの1行目がJに入る
    int_list = [ int(k) for k in j ]#jの中身をｋ
    ave += np.mean( int_list )#シートの枚数分回して合計の平均値
ave_fin = ave/ len( list_len )
print(ave_fin)
