import xlwings as xw
from itertools import pairwise
from dataclasses import dataclass

@dataclass
class Record:
    key: int
    category: str
    title: str
    input: float
    output: float
    remark: str

wb = xw.Book( 'sample.xlsx' )
name = wb.names[ 'sample' ]

# 値のある範囲に絞る
range = name.refers_to_range.current_region

# ヘッダ以外を二次元配列として取得
rows = range.value[1:]

# レコードの先頭行のインデックスの配列を作る
recordIndex = [ index for [ index, row ] in enumerate( rows ) if row[0] is not None ]
recordIndex.append( len( rows) )

# レコードの配列を作成
records: list[ Record ] = []
for [ start, end ] in pairwise( recordIndex ):
    # 開始行～終了行(含まない)でスライス
    data = rows[ start:end ]

    # スライスしたデータからRecordを作成    
    records.append( Record(
        key = data[0][0],
        category = data[0][1],
        title = data[0][2],
        input = data[0][4],
        output = data[1][4],
        remark = data[2][4]
    ) )

for record in records:
    print( record )
