import ole
import hexdump
import sys
import struct
import about_property
"""
if len(sys.argv) != 4:
    print('input value error!')
    exit()
"""
fp = open('macro1.hwp', 'rb')


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

def get_all_small_block():
    return 0

def get_entry_list(bat, start_entry) :
    cluster = start_entry
    cluster_list = [cluster]
    while True:
        cluster_bytes =  bat[cluster * 4 : (cluster + 1) * 4]
        cluster = struct.unpack('<I', cluster_bytes)[0]
        if cluster == 0xfffffffe:
            break
        
        cluster_list.append(cluster)
    return cluster_list

#cluster_list = get_cluster_list(2, bbat) ## bbat에서 클러스터 리스트 얻기

def get_all_property(bbat, start_entry_of_property):
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
bbat = get_all_block(header['array_bbat'])
entry_list = get_entry_list(bbat, header['start_entry_of_property'])

property_data = get_all_property(bbat, header['start_entry_of_property'])

for i in range(13):
    property_jscript_info = get_property_info(property_data, i)
    print(property_jscript_info)    

"""
def get_sbat(start_cluster_of_sbat, number_sbat_depot):
    global bbat
    cluster_list = get_clusters(start_cluster_of_sbat, bbat)
    print("cluster_list:", cluster_list)
    return

print_info()
get_sbat(start_cluster_of_sbat, number_sbat_depot)




data[:filesize] 이런식으로 처리하면 클러스터 크기 만큼이 아닌 정확한 크기만큼 가져올 수 있다.

"""

fp.close()