import ole
import hexdump
import sys
import struct

if len(sys.argv) != 4:
    print('input value error!')
    exit()

fp = open('macro1.hwp', 'rb')
Buffer = ole.ReadBlock(fp, int(sys.argv[1]))


hexdump.Buffer(Buffer, int(sys.argv[2]), int(sys.argv[3]))


header_block = ole.ReadBlock(fp, -1)
# 헤더를 분석하는 파일을 만들자!
magic_number = hexdump.Dump(header_block, 0, 8) # d0 cf 11 e0 a1 b1 1a e1

print(magic_number)

number_bbat_depot = struct.unpack('<I',hexdump.Dump(header_block, 44, 4))[0]
start_block_of_property = struct.unpack('<I', hexdump.Dump(header_block, 48, 4))[0]
start_block_of_sbat = struct.unpack('<I', hexdump.Dump(header_block, 60 ,4))[0]
number_sbat_depot = struct.unpack('<I', hexdump.Dump(header_block, 64, 4))[0]

iter_bbat = struct.iter_unpack('<I', hexdump.Dump(header_block, 76, 4 * number_bbat_depot))
array_bbat = []
for i in range(0, number_bbat_depot):
    array_bbat.append(next(iter_bbat)[0])
print(array_bbat)

bbat = b""
for idx in array_bbat:
   bbat += hexdump.Dump(ole.ReadBlock(fp, idx), 0)


def get_clusters(cluster, bbat) :
    cluster_list = [cluster]
    while True:
        cluster_bytes =  bbat[cluster * 4 : (cluster + 1) * 4]
        cluster = struct.unpack('<I', cluster_bytes)[0]
        if cluster == 0xfffffffe:
            break
        
        cluster_list.append(cluster)
    return cluster_list

cluster_list = get_clusters(2, bbat)

print(cluster_list)




fp.close()