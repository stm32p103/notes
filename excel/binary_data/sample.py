import xlwings as xw
from itertools import pairwise
from dataclasses import dataclass

@dataclass
class DataA:
    x: float
    y: float

class DataB:
    x: float
    y: float

def updateFromA( text: str ):
    pass

def updateFromB( text: str ):
    pass

class Data:
    data: list[ int ]

@dataclass
class Field:
    startByte: int
    startBit: int
    length: int
    isBigEndian: bool

class Byte:
    def __init__( self, value: int ):
        if value < 0 or value > 0xFF:
            raise ValueError( 'value must be 0-255' )
        self._value = value

    @property
    def value( self ):
        return self._value

from math import floor
class MultibyteData:
    def __init__( self, data: list[ int ] ):
        self._data = data[:]

    def __getitem__( self, index: int ):
        return self._data[ index ]
    
    def slice( self, field: Field ):
        # 取り出した結果の格納先
        sliced: list[ int ] = []
        totalBytes = floor( field.length / 8 )

        # 下位・上位ビットのデータを取り出すマスク
        currentMask = ( 0xFF << field.startBit ) & 0xFF
        nextMask = ( 0xFF >> ( 8 - field.startBit ) ) & 0xFF

        if field.isBigEndian:
            step = -1
        else:
            step = 1

        # フルにある
        for i in range( field.startByte, field.startByte + step * totalBytes, step ):
            currentValue = ( data[ i ] & currentMask ) >> field.startBit            
            nextValue = ( data[ i + step ] & nextMask ) << ( 8 - field.startBit )
            value = currentValue + nextValue
            sliced.append( value )

        # 次のステップに進む
        i = i + step

        # 部分的にある
        residual = field.length % 8
        if residual > 0:
            if residual + field.startBit <= 8:
                mask = 0xFF & ( 0xFF >> ( 8 - residual ) )
                value = ( data[ i ] >> field.startBit ) & mask
            else:
                mask = 0xFF & ( 0xFF >> ( residual + field.startBit - 8 ) )
                currentValue = ( data[ i ] & currentMask ) >> field.startBit            
                nextValue = ( data[ i + step ] & mask ) << ( 8 - field.startBit )
                value = currentValue + nextValue
            sliced.append( value )

        # 部分的にある
        residualMask = ( 0xFF >> ( 8 - residual ) ) & 0xFF
        sliced[ -1 ] = sliced[ -1 ] & residualMask

        return MultibyteData( sliced )


# def test( data: list[ int ] ):
#     startBit = 5
#     startByte = 6
#     length = 42

#     totalBytes = floor( length / 8 )

#     # 下位・上位ビットのデータを取り出すマスク
#     lowerMask = ( 0xFF << startBit ) & 0xFF
#     upperMask = ( 0xFF >> ( 8 - startBit ) ) & 0xFF

#     # 取り出した結果の格納先
#     sliced: list[ int ] = []
#     for i in range( startByte, startByte - totalBytes - 1, -1 ):
#         lowerValue = ( data[ i ] & lowerMask ) >> startBit
#         upperValue = ( data[ i - 1 ] & upperMask ) << ( 8 - startBit )
#         value = lowerValue + upperValue
#         sliced.append( value )
    
#     residual = length % 8
#     residualMask = ( 0xFF >> ( 8 - residual ) ) & 0xFF
#     sliced[ -1 ] = sliced[ -1 ] & residualMask

#     print( ' '.join( [ f'{byte:08b}' for byte in sliced ] ) )

print( 'Big endian' )
for i in range( 0, 56 ):
    data = [ 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 ]
    byte = floor( i / 8 )
    bit = i % 8

    data[ byte ] = 0x80 >> bit
    mb = MultibyteData( data )
    f = Field( 6, 3, 42, True )
    s = mb.slice(f)
    print( ' '.join( [ f'{byte:08b}' for byte in s._data ] ) )


print( 'Little endian' )
for i in range( 0, 56 ):
    data = [ 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 ]
    byte = floor( i / 8 )
    bit = i % 8

    data[ byte ] = 0x01 << bit
    mb = MultibyteData( data )
    f = Field( 1, 3, 42, False )
    s = mb.slice(f)
    print( ' '.join( [ f'{byte:08b}' for byte in s._data ] ) )
