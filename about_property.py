import struct 

def get_type(property_data):
    property_type = struct.unpack('<B',property_data[66: 66 + 1])[0]
    if property_type == 1:
        return 'storage'
    elif property_type == 2:
        return 'stream'
    elif property_type == 5:
        return 'root'


def get_starting_block_of_property(property_data): # 블록 내부에서 프로퍼티
    start_block_of_property = struct.unpack('<I', property_data[116 : 120])[0]
    return start_block_of_property


def get_size_of_property(property_data):
    size_of_property = struct.unpack('<I', property_data[120:124])[0]
    return size_of_property # 만약 0x1000보다 크면 BBAT, 작다면 SBAT