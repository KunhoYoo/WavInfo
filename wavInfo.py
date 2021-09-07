# This Python file uses the following encoding: utf-8
# -*- coding: utf-8 -*-

# Dependencies: ffmpeg
#               http://www.ffmpeg.org/

import os
from pydub import AudioSegment
from pydub.utils import mediainfo
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.append(['파일명', '샘플링레이트', '채널', '음질', '지속시간', '샘플의개수', '디스크 사이즈', '최고진폭'])

file_list = [f for f in os.listdir('./') if f.endswith('.m4a') or f.endswith('.wav')]
for file_name in file_list:
    info1 = AudioSegment.from_file(file_name)
    info2 = mediainfo(file_name)

    row = []
    row.append(info2['filename'])
    row.append(info2['sample_rate'] + ' Hz')
    row.append('mono' if info2['channels'] == '1' else 'stereo')
    row.append('{} bits'.format(info1.sample_width * 8))
    row.append('{:02d}:{:02.03f}'.format(int(float(info2['duration']) // 60), float(info2['duration']) % 60))
    row.append(int(int(info2['sample_rate']) * float(info2['duration'])))
    row.append('{:.02f} kB'.format(int(info2['size']) / 1024))
    row.append('{:.2f} dB'.format(info1.max_dBFS))
    
    ws.append(row)
wb.save('result.xlsx')
