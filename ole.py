def ReadBlock(fp, block_num) :
    block_buf = ""

    # -1 블록의 위치가 0이며, 0 블록이 0x200 위치가 되도록
    # 아래의 공식을 사용한다.
    fp_pos = (block_num + 1) * 0x200

    try: 
        fp.seek(fp_pos)
        block_buf = fp.read(0x200)
    except:
        pass

    return block_buf # 읽은 내용을 리턴한다.

