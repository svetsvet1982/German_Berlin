"""기획서 '화별 개요' H열의 작성 상태를 바꾼다: python mark_done.py 기획서.xlsx 화번호 [상태]"""
import sys, openpyxl
path, ep = sys.argv[1], int(sys.argv[2]); st = sys.argv[3] if len(sys.argv) > 3 else '초고완료'
wb = openpyxl.load_workbook(path); ws = wb['화별 개요']
assert ws.cell(ep + 1, 1).value == ep
ws.cell(ep + 1, 8).value = st; wb.save(path); print(ep, ws.cell(ep + 1, 2).value, '→', st)
