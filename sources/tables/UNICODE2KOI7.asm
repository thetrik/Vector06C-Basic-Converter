
; FASM
; Unicode to VECTOR-06C KOI7 translation table

; Fill ascending values at specified offset
macro fill_inc initval, start, count {
    repeat count
	store byte % - 1 + initval at % - 1 + start
    end repeat
}

; Put values at specified offset
macro put start, [args] {
common
    local c
    c = 0
forward
    store byte args at start + c
    c = c + 1
}

; Put value at specified offsets
macro put_multiple value, [start*] {
forward
    store byte value at start
}

; Fill all range with spaces
repeat 65536
    db ' '
end repeat

; Modify needed values
fill_inc 0, 0, 0x60
put 0x60, ' '
fill_inc 'A', 'a', 26
put 0x7b, ' ', ' ', ' ', ' ', 0x7f
put 0xa4, '$'

; Cyrillic
put 0x410, 'a', 'b', 'w', 'g', 'd', 'e', 'v', 'z', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', \
	   'r', 's', 't', 'u', 'f', 'h', 'c', '~', '{', '}', 'x', 'y', 'x', '|', '`', 'q'
put 0x263a, 1, 2  ; Smiles
put_multiple 3, 0x2665, 0x2764, 0x2661	; Heart
put_multiple 4, 0x25C6, 0x2666, 0x25C7, 0x2662, 0x2BC1	; Diamond
put_multiple 5, 0x2663, 0x2667	 ; Club suit
put_multiple 6, 0x2660, 0x2664	 ; Spade suit
put_multiple 7, 0x2669, 0x266A	 ; Note
put 0x25d8, 8
put_multiple 9, 0x25AA, 0x25AB	  ; Rect
put_multiple 11, 0x260C, 0x2642  ; Conjunction/male
put 0x2640, 12	; Female
put_multiple  14, 0x266C, 0x266B  ; Beamed Sixteenth Notes
put_multiple 15, 0x263C, 0x2638, 0x2699, 0x26ED ; Sun / Gear
put_multiple 16, 0x23F5, 0x25B6, 0x25B7, 0x25B8, 0x25B9, 0x2BC8 ; Right-Pointing Triangle
put_multiple 17, 0x23F4, 0x25C0, 0x25C1, 0x25C2, 0x25C3, 0x2BC7 ; Left-Pointing Triangle
put_multiple 18, 0x2195, 0x21D5, 0x21F3, 0x2B0D, 0x2B65 ; Up Down Arrow
put 0x203c, 19 ; Double Exclamation Mark
put_multiple 20, 0x3C0, 0x213C	 ; PI
put 0xa7, 21	  ; Section
put_multiple 22, 0x2582, 0x2583, 0x2584 ; Lower block
put_multiple 23, 0x21DE, 0x21EF ; Upwards Arrow with Double Stroke
put_multiple 24, 0x2191, 0x21d1, 0x21e7, 0x2b06, 0x2b61 ; Up Arrow
put_multiple 25, 0x2193, 0x21d3, 0x21e9, 0x2b07, 0x2b63 ; Down Arrow
put_multiple 26, 0x2192, 0x21d2, 0x21e8, 0x2b95, 0x2b62 ; Right Arrow
put_multiple 27, 0x2190, 0x21d0, 0x21e6, 0x2b05, 0x2b60 ; Left Arrow
put 0x02fe, 28 ; Righthand Interior Product
put_multiple 29, 0x20E1, 0x2194, 0x21D4, 0x21FF, 0x27F7, 0x27FA, 0x2B0C ; Left Right Arrow
put_multiple 30, 0x23f6, 0x25b2, 0x25b3, 0x25b4, 0x25b5, 0x2bc5, 0x2616, 0x2617  ; Shogi Piece / Up Triangle
put_multiple 31, 0x23f7, 0x25bc, 0x25bd, 0x25be, 0x25bf, 0x2bc6  ; Down Triangle
put_multiple 127, 0x2B1B, 0x2BC0, 0x25AA, 0x25a0, 0x25FE, 0x2588, 0x2589 ;Black Square