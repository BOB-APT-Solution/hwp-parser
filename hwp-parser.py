import ole
import hexdump
import sys
import struct
import about_property
import zlib

fp = open(sys.argv[1], 'rb')


#Buffer = ole.ReadBlock(fp, int(sys.argv[1]))
#hexdump.Buffer(Buffer, int(sys.argv[2]), int(sys.argv[3]))

def get_header_info():
    header = {}
    header_block = ole.ReadBlock(fp, -1)
    header['magic_number'] = hexdump.Dump(header_block, 0, 8) # d0 cf 11 e0 a1 b1 1a e1
    header['number_bbat_depot'] = struct.unpack('<I',hexdump.Dump(header_block, 44, 4))[0]
    header['start_entry_of_property'] = struct.unpack('<I', hexdump.Dump(header_block, 48, 4))[0]
    header['start_cluster_of_sbat'] = struct.unpack('<I', hexdump.Dump(header_block, 60 ,4))[0]
    header['number_sbat_depot'] = struct.unpack('<I', hexdump.Dump(header_block, 64, 4))[0]
    iter_bbat = struct.iter_unpack('<I', hexdump.Dump(header_block, 76, 4 * header['number_bbat_depot']))
    header['array_bbat'] = []
    for i in range(0, header['number_bbat_depot']):
        header['array_bbat'].append(next(iter_bbat)[0])
    return header

def print_info(header):
    print('magic number:', header['magic_number'])
    print("number_bbat_depot:", header["number_bbat_depot"])
    print('start_entry_of_property:', header['start_entry_of_property'])
    print("start_cluster_of_sbat:", header['start_cluster_of_sbat'])
    print("number_sbat_depot:", header['number_sbat_depot'])
    print("array_bbat:",header['array_bbat'])

def get_all_block(entry_list):
    blocks = b""
    for idx in entry_list:
        blocks += hexdump.Dump(ole.ReadBlock(fp, idx), 0)

    return blocks

def get_all_small_block(small_blocks, entry_list):
    blocks = b""
    for idx in entry_list:
        blocks += small_blocks[idx * 0x40: (idx + 1) * 0x40]

    return blocks


def get_entry_list(bat, start_entry) : # 일반적인 entry list를 얻을 때 사용
    cluster = start_entry
    cluster_list = [cluster]
    while True:
        cluster_bytes =  bat[cluster * 4 : (cluster + 1) * 4]
        cluster = struct.unpack('<I', cluster_bytes)[0]
        if cluster == 0xfffffffe:
            break
        
        cluster_list.append(cluster)
    return cluster_list

def get_all_property(bbat, start_entry_of_property): # property에 관한 entry list와 blocks을 얻을 수 있음
    property_entry_list = get_entry_list(bbat, start_entry_of_property)
    property_blocks = get_all_block(property_entry_list)
    print("property_entry_list:", property_entry_list)
    return property_blocks

def get_property_info(property_blocks, index): 
    """
    0: root entry
    1: file header
    2: doc info
    3: hwp summary information
    4: body text
    5: prv image
    6: prv text
    7: doc options
    8: scripts
    9: jscript version
    10: default jscript
    11: _link doc
    파일에 따라 조금씩 다를 수 있지만, 11까진 거의 확정(확인 필요)
    """
    property_data = hexdump.Dump(property_blocks, index * 0x80, 0x80)
    dic_property = {
        'len_name' : struct.unpack('<H', property_data[64: 66])[0],
    }
    dic_property['name'] = property_data[: dic_property['len_name'] - 2].decode('utf-16')
    dic_property['type'] = about_property.get_type(property_data)
    dic_property['start_block'] = about_property.get_starting_block_of_property(property_data)
    dic_property['size'] = about_property.get_size_of_property(property_data)

    return dic_property

header = get_header_info()
print_info(header)

bbat = get_all_block(header['array_bbat']) # bbat 데이터
#entry_list = get_entry_list(bbat, header['start_entry_of_property']) # 프로퍼티 entry(클러스터) 리스트 get_all_property 안에서 동작
sbat_entry_list = get_entry_list(bbat, header['start_cluster_of_sbat']) # sbat entry 리스트

sbat = get_all_block(sbat_entry_list) # sbat 데이터

property_data = get_all_property(bbat, header['start_entry_of_property']) # property data

# property info
property_jscript_info = get_property_info(property_data, 10)
print("JScript property info:",property_jscript_info)

if property_jscript_info['size'] > 4096:
    big_data = get_all_block(get_entry_list(bbat, property_jscript_info['start_block']))[:property_jscript_info['size']]
    decompressed_data = zlib.decompress(big_data, -15)
else:
    property_root_info = get_property_info(property_data, 0)
    js_sbat_entry_list = get_entry_list(sbat, property_jscript_info['start_block']) # sbat entry list
    small_data_block_list = get_entry_list(bbat, property_root_info['start_block'])
    small_data_blocks = get_all_block(small_data_block_list)
    data = get_all_small_block(small_data_blocks, js_sbat_entry_list)[:property_jscript_info['size']]
    decompressed_data = zlib.decompress(data, -15)

print("========================= JavaScript ===============================")

result_script = decompressed_data.decode('utf-16')

import os
print(result_script)
result_path = os.path.join(os.path.abspath(sys.argv[1]), os.path.pardir, f"{sys.argv[1].split('.')[0]}.js")
with open(result_path, 'w') as result:
    result.write(result_script)

fp.close()

"""

data[:filesize] 이런식으로 처리하면 클러스터 크기 만큼이 아닌 정확한 크기만큼 가져올 수 있다.

"""
