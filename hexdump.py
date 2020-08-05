import sys

def File( fname, start, size=0x200, width=16 ):
     fp = open( fname, "rb" )
     fp.seek(start)
     row = start % width           # 열
     col = (start / width) * width # 행
     r_size = 0
     line_start = row
     while True :
          if (r_size + (width-line_start) < size) :
               r_char = (width-line_start) # 읽어야할 문자 수
               r_size += (width-line_start)
          else :
               r_char = size - r_size
               r_size = size
           
          # print line_start, r_char
          line = fp.read(r_char)
          if len(line) == 0 :
               break
          # 주소 값
          output = "{:08x} : ".format(int(col))
          # Hex 값     
          for c in line:  
            output += line_start * "   " \
               + "".join( "{:02x}".format(c))
          output += "  " \
                + (width - (line_start + r_char) ) * "   "
          # 문자 값
          output += line_start * " "
     
          output += "".join([chr(c) if IsPrint(c) else '.' for c in line])
          print(output)
          col += width
          line_start = 0
          if r_size == size :
               break
     fp.close()

def Buffer( buf, start, size=0x200, width=16 ):
     # 주어진 버퍼의 크기가 size보다 작다면 size값을 조정
     if len ( buf ) < size :
         size = len ( buf )
     row = start % width    # 열
     col = (start // width)  # 행
     # [row ... width*col]
     # [width*col ... width * (col+1)]
     r_size = 0
     line_start = row + ( col * width )
     # print hex(line_start), hex(width*(col+1))
     # print hex(row), hex(col)

     while True :
          line = buf[line_start:width * (col + 1)]
           
          if len(line) == 0 :
               break
          if ( ( r_size + len ( line ) ) < size ) :
               pass
          else :
               #print hex(line_start), hex(line_start + (size - r_size))
               line = line[0:(size - r_size)]
               r_size = size - len ( line )
          # 주소 값
          output = "{:08x}: ".format( int(( line_start / width ) * width) )
          # Hex 값   
          for c in line:   
            output += row * " " \
               + "".join( "{:02x} ".format(c))
          output += "  " \
               + (width - (row + len(line)) ) * " "
          # 문자 값
          output += row * " "
          output += "".join( [chr(c) if IsPrint(c) else '.' for c in line] )
          print(output)
          line_start = width * ( col + 1 )
          col += 1
          row = 0
          r_size += len( line )
          if r_size == size :
               break

def Dump ( buf, start, size=0x200):
     
     return buf[start: start+ size]

def IsPrint ( char ) :
     c = ord(chr(char))
     if c >= 0x20 and c < 0x80 :
          return True
     else :
          return False

if __name__ == '__main__':
     File( sys.argv[1], 
          int( sys.argv[2] ),
          int( sys.argv[3] ),
          int( sys.argv[4] ))